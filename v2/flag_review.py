"""Run-bound owner Review Decisions. Parser evidence is immutable in this layer."""
from __future__ import annotations
import hashlib,json
from pathlib import Path
from nmdc_profiler.update_engine import _write_json_atomic
from nmdc_profiler.review_decisions import ALLOWED_DECISIONS,ALLOWED_STATUSES,_write_parser_mapping_fix_request

def fingerprint(value):
    return hashlib.sha256(json.dumps(value,sort_keys=True,ensure_ascii=False,separators=(",",":"),default=str).encode()).hexdigest()

class FlagReview:
    def __init__(self,stage,rows,baseline):
        self.stage=stage
        self.rows=[dict(r) for r in rows]
        self.raw=list(stage["flags"])
        self.ids=[fingerprint([i,r]) for i,r in enumerate(self.raw)]
        self.path=stage["stage_dir"]/"flag_review.json"
        self.identity=fingerprint([stage["run_id"],baseline,self.ids])
        self.payload=json.loads(self.path.read_text(encoding="utf-8")) if self.path.exists() else {"identity":self.identity,"version":0,"dirty":False,"choices":{}}
        if self.payload["identity"]!=self.identity:raise ValueError("Flag evidence changed; restage before editing.")
        for id,r in zip(self.ids,self.rows):
            r["Flag ID"]=id
            c=self.payload["choices"].get(id)
            if c:r.update({"User Decision":c["decision"],"User Comment":c["comment"],"Resolution Status":c["status"]})
    def edit(self,body,save):
        if body.get("run")!=self.stage["run_id"] or body.get("flag_version")!=self.payload["version"]:raise ValueError("Flag review changed; reload.")
        changes=body.get("changes")
        if not isinstance(changes,list) or len(changes)>len(self.ids):raise ValueError("Invalid flag changes")
        choices=dict(self.payload["choices"]);seen=set()
        for c in changes:
            if not isinstance(c,dict) or c.get("id") not in self.ids or c["id"] in seen:raise ValueError("Unknown or duplicate flag identity")
            decision=c.get("decision","");status=c.get("status","OPEN");comment=c.get("comment","")
            if decision not in ALLOWED_DECISIONS|{""} or status not in ALLOWED_STATUSES or not isinstance(comment,str) or len(comment)>10000:raise ValueError("Invalid flag decision")
            seen.add(c["id"]);choices[c["id"]]={"decision":decision,"status":status,"comment":comment}
        self.payload.update(choices=choices,version=self.payload["version"]+1,dirty=not save)
        _write_json_atomic(self.path,self.payload)
        return {"flag_version":self.payload["version"],"flag_dirty":self.payload["dirty"]}
    def report(self,body,parent,state):
        if self.payload["dirty"]:raise ValueError("Save flag decisions before creating a report.")
        ids=body.get("ids");comment=body.get("comment","")
        if not isinstance(ids,list) or not ids or len(set(ids))!=len(ids) or any(i not in self.ids for i in ids):raise ValueError("Invalid flag selection")
        changes=[{"id":i,"decision":"NEEDS PARSER/MAPPING FIX","status":"OPEN","comment":comment} for i in ids]
        self.edit(dict(body,changes=changes),True)
        flags=[dict(self.raw[self.ids.index(i)],user_decision="NEEDS PARSER/MAPPING FIX",resolution_status="OPEN",user_comment=comment) for i in ids]
        path=_write_parser_mapping_fix_request(state,self.stage["run_id"],flags,request_dir=parent)
        return {"run_id":self.stage["run_id"],"parser_mapping_fix_requests":len(flags),"fix_request_file":str(path),"fix_request_folder":str(path.parent),"request_sequence":path.parent.name,"flag_version":self.payload["version"]}
