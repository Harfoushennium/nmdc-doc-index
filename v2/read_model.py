"""Disposable SQLite read model. Never authoritative; rebuilt from approved views."""
from __future__ import annotations
import json, re, sqlite3
from collections import Counter
def tokens(query):
    out=[]
    for m in re.finditer(r'(-?)"([^"]+)"|([^\s+]+)',query):
        if m.group(2) is not None:out.append((bool(m.group(1)),True,m.group(2).casefold()))
        elif m.group(3)!="-":out.append((m.group(3).startswith("-"),False,m.group(3).lstrip("-").casefold()))
    return out
class ReadModel:
    # In-memory SQLite: snapshot-scoped, disposable, no live database in OneDrive.
    def __init__(self,rows,fields):
        self.fields=list(fields)
        self.rows=rows
        self.db=sqlite3.connect(":memory:",check_same_thread=False)
        self.db.execute("CREATE TABLE records (position INTEGER PRIMARY KEY, all_text TEXT NOT NULL, "+",".join(f"c{i} TEXT NOT NULL" for i in range(len(fields)))+")")
        self.db.executemany("INSERT INTO records VALUES ("+",".join("?" for _ in range(len(fields)+2))+")",
            ((i," ".join(str(v).casefold() for v in r.values()),*(str(r.get(k,"")).casefold() for k in fields)) for i,r in enumerate(rows)))
        for k in ("Document No.","Project No.","Revision","Source File","Event_Key"):
            if k in fields:
                col=self.col(k);self.db.execute(f"CREATE INDEX idx_{col} ON records({col})")
        self.db.commit()
    def col(self,name):
        if name not in self.fields:raise ValueError("Unknown index field")
        return "c"+str(self.fields.index(name))
    def predicate(self,query,column,filters):
        clauses=[];args=[]
        for negative,exact,value in tokens(query):
            if column:
                c=self.col(column)
                clause=f"{c}=?" if exact else f"instr({c},?)>0"
                # For missing values, the old search returns empty; indexed column does too.
                args.append(value)
            elif exact:
                # All-column exact is evaluated against the actual record values
                # (not synthesized missing fields), using the existing Python fallback.
                raise NotImplementedError
            else:
                clause="instr(all_text,?)>0";args.append(value)
            clauses.append(("NOT ("+clause+")") if negative else clause)
        if not isinstance(filters,dict):raise ValueError("Invalid filters")
        for name,values in filters.items():
            c=self.col(name)
            if not isinstance(values,list) or len(values)>500 or any(not isinstance(x,str) for x in values):raise ValueError("Invalid filter values")
            if values:
                clauses.append(c+" IN ("+",".join("?" for _ in values)+")")
                args.extend(x.casefold() for x in values)
        return (" WHERE "+" AND ".join(clauses) if clauses else ""),args
    def page(self,query="",column="",sort="",desc=False,filters=None,page=0,group=""):
        where,args=self.predicate(query,column,filters or {})
        count=self.db.execute("SELECT count(*) FROM records"+where,args).fetchone()[0]
        order=(" ORDER BY "+self.col(sort)+(" DESC" if desc else " ASC")+",position ASC") if sort else " ORDER BY position"
        ids=self.db.execute("SELECT position FROM records"+where+order+" LIMIT 100 OFFSET ?",[*args,page*100]).fetchall()
        groups=[]
        if group:
            c=self.col(group)
            groups=[{"value":self.rows[pos].get(group,""),"count":n} for pos,n in self.db.execute(f"SELECT min(position),count(*) FROM records{where} GROUP BY {c} ORDER BY count(*) DESC,{c} LIMIT 100",args)]
        return {"count":count,"fields":self.fields,"rows":[self.rows[i[0]] for i in ids],"groups":groups}
    def close(self):self.db.close()

def apply_filters(rows,filters,fields):
    if not isinstance(filters,dict):raise ValueError("Invalid filters")
    for k,values in filters.items():
        if k not in fields or not isinstance(values,list) or len(values)>500 or any(not isinstance(x,str) for x in values):raise ValueError("Invalid column filter")
        if values:
            allowed={x.casefold() for x in values}
            rows=[r for r in rows if str(r.get(k,"")).casefold() in allowed]
    return rows
