"""Derived custom columns, compatible with the legacy CustomFields/KeywordMappings tables.
No canonical event or approved version is modified."""
from __future__ import annotations
import fnmatch, re
from collections import defaultdict
FIELD_COLUMNS=("Enabled?","Field Name","Search In","Match Behavior","Separator","Notes")
MAP_COLUMNS=("Enabled?","Field Name","Keyword / Pattern","Result","Match Type","Priority","Notes")
DOC_COLUMNS=("Flag Level","Project No.","Source Family","Discipline","Category","Subcategory","Document No.","Document Title","Company Document No.","Latest Revision","Latest Event Date","Latest Event Status","Document Link","Source File","Source Sheet","Source Row","Source Cell","Global Document Key","Original_Hyperlink_Target","Source Workbook Path")
def enabled(v):
    if isinstance(v,bool):return v
    value=str(v).strip().upper()
    if value in {"TRUE","YES","Y","1","ON","CHECKED"}:return True
    if value in {"FALSE","NO","N","0","OFF","UNCHECKED",""}:return False
    raise ValueError("Invalid Enabled? value")
def validate(payload,reserved):
    if not isinstance(payload,dict) or set(payload)!={"fields","mappings"}:raise ValueError("Invalid custom field configuration schema")
    fields,maps=payload.get("fields"),payload.get("mappings")
    if not isinstance(fields,list) or not isinstance(maps,list) or len(fields)>200 or len(maps)>5000:raise ValueError("Invalid custom-field table size")
    protected={k.casefold() for k in (*DOC_COLUMNS,*reserved)}
    names=set();clean=[];out=[]
    for f in fields:
        if not isinstance(f,dict) or set(f)!=set(FIELD_COLUMNS):raise ValueError("Invalid CustomFields schema")
        if any(not isinstance(v,(str,bool)) for v in f.values()):raise ValueError("Invalid field value")
        f=dict(f);name=str(f["Field Name"]).strip();key=name.casefold()
        if not name or len(name)>100 or key in protected or key in names or key.startswith("__nmdc_"):raise ValueError("Reserved, missing or duplicate custom field: "+name)
        f["Field Name"]=name;f["Enabled?"]=enabled(f["Enabled?"])
        f["Match Behavior"]=str(f["Match Behavior"]).strip().upper() or "ALL UNIQUE"
        if f["Match Behavior"] not in ("FIRST","ALL UNIQUE"):raise ValueError("Unknown match behavior")
        f["Search In"]=str(f["Search In"]).strip() or "Document Title"
        f["Separator"]=str(f["Separator"]) or " | "
        clean.append(f);names.add(key)
    available={k.casefold() for k in DOC_COLUMNS}|names
    field_order={f["Field Name"].casefold():i for i,f in enumerate(clean)}
    for i,f in enumerate(clean):
        search=f["Search In"]
        if search.upper() not in ("ALL TEXT","*ALL TEXT*"):
            cols=[k.strip().casefold() for k in re.split("[;,]",search) if k.strip()]
            if not cols or not set(cols)<=available or f["Field Name"].casefold() in cols:raise ValueError("Unknown/self-referencing search column")
            dependencies=[field_order[c] for c in cols if c in field_order]
            if any(dep>=i for dep in dependencies):raise ValueError("Custom field dependency must refer to a prior field")
    for m in maps:
        if not isinstance(m,dict) or set(m)!=set(MAP_COLUMNS):raise ValueError("Invalid KeywordMappings schema")
        if any(not isinstance(v,(str,bool,int,float)) for v in m.values()):raise ValueError("Invalid mapping value")
        m=dict(m);m["Enabled?"]=enabled(m["Enabled?"])
        m["Field Name"]=str(m["Field Name"]).strip()
        m["Keyword / Pattern"]=str(m["Keyword / Pattern"]).strip()
        m["Result"]=str(m["Result"]);m["Match Type"]=str(m["Match Type"]).strip().upper() or "CONTAINS"
        if m["Field Name"].casefold() not in names or not m["Keyword / Pattern"] or len(m["Keyword / Pattern"])>512:raise ValueError("Unknown field or empty/oversize keyword")
        if m["Match Type"] not in ("CONTAINS","ALL TERMS","EXACT","WILDCARD"):raise ValueError("Unknown keyword match type")
        try:m["Priority"]=str(round(float(m["Priority"]))) if str(m["Priority"]).strip() else "999999"
        except (ValueError,OverflowError):raise ValueError("Invalid keyword priority")
        # Unsupported VBA Like constructs are never silently approximated.
        if m["Match Type"]=="WILDCARD" and any(ch in m["Keyword / Pattern"] for ch in "[]#"):raise ValueError("Legacy wildcard contains unsupported VBA Like constructs; review mapping")
        out.append(m)
    return {"fields":clean,"mappings":out}
def wildcard(text,pattern):
    stars=pattern.count("*");questions=pattern.count("?")
    if stars==1:
        pre,suf=pattern.split("*")
        start=0
        if pre:
            match=re.search(re.escape(pre),text,re.IGNORECASE)
            if not match:return False,""
            start=match.end()
        end=len(text)
        if suf:
            match=re.search(re.escape(suf),text[start:],re.IGNORECASE)
            if not match:return False,""
            end=start+match.start()
        return True,text[start:end]
    if questions:
        plain=pattern.replace("?","")
        if plain and (pattern.endswith("?"*questions) or pattern.startswith("?"*questions)):
            match=re.search(re.escape(plain),text,re.IGNORECASE)
            if match:
                if pattern.endswith("?"*questions) and match.end()+questions<=len(text):return True,text[match.end():match.end()+questions]
                if pattern.startswith("?"*questions) and match.start()>=questions:return True,text[match.start()-questions:match.start()]
    return fnmatch.fnmatchcase(text.casefold(),"*"+pattern.casefold()+"*"),""
def match_value(text,m):
    pattern=m["Keyword / Pattern"];kind=m["Match Type"];fold=text.casefold();term=pattern.casefold();extract=""
    if kind=="EXACT":ok=fold.strip()==term
    elif kind=="ALL TERMS":
        ok=all((t[1:] not in fold if t.startswith("-") and len(t)>1 else t in fold) for t in term.replace("+"," ").split(" ") if t)
    elif kind=="WILDCARD":ok,extract=wildcard(text,pattern)
    else:ok=term in fold
    if not ok:return None
    result=m["Result"]
    if not result:return (extract or pattern).strip()
    if extract:
        if "*" in result:return result.replace("*",extract).strip()
        needle="?"*len(extract)
        if needle in result:return result.replace(needle,extract).strip()
    return result.strip()
def enrich(rows,payload):
    fields=payload["fields"];groups=defaultdict(list)
    for m in payload["mappings"]:
        if enabled(m["Enabled?"]):groups[m["Field Name"].casefold()].append(m)
    for mappings in groups.values():mappings.sort(key=lambda m:int(m["Priority"]))
    for row in rows:
        # Materialize fields first, then apply sequentially, retaining V1 field ordering.
        for f in fields:row[f["Field Name"]]=""
        for f in fields:
            name=f["Field Name"]
            if not enabled(f["Enabled?"]):continue
            search=f["Search In"]
            if search.upper() in ("ALL TEXT","*ALL TEXT*"):
                cols=[k for k in row if k!=name and not k.upper().startswith("__NMDC_")]
            else:
                by_name={k.casefold():k for k in row}
                cols=[by_name[k.strip().casefold()] for k in re.split("[;,]",search) if k.strip().casefold() in by_name]
            text=" | ".join(str(row[k]) for k in cols if row.get(k) not in (None,""))
            if not text:continue
            values=[];seen=set()
            for m in groups[name.casefold()]:
                v=match_value(text,m)
                if v is None:continue
                if f["Match Behavior"]=="FIRST":values=[v];break
                if v and v.casefold() not in seen:seen.add(v.casefold());values.append(v)
            row[name]=f["Separator"].join(values)
    return rows
