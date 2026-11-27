"""NMDC V2 standalone integration candidate. Existing engine owns approved data."""
from __future__ import annotations
import argparse, csv, hashlib, hmac, io, json, os, secrets, shlex, shutil, sys, threading, time, webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit, parse_qs
import xlsxwriter
from nmdc_profiler.excel_bridge import load_approved_state, load_latest_stage, build_document_rows, build_revision_rows, build_flag_rows
from nmdc_profiler.extractor import OUTPUT_FIELDS
from nmdc_profiler.file_review import affected_files, save_file_review, validated_review
from nmdc_profiler.update_engine import approve_stage, reject_stage, hold_stage, _write_json_atomic, stage_update
from nmdc_profiler.runtime_engine import stage_runtime_update
from nmdc_profiler.runtime_admin import undo_last_approval
from nmdc_profiler.source_access import install_resilient_source_access
from nmdc_profiler.layout_compat import install_layout_compatibility

ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
WEB = ROOT / "v2"
CORE = {"events": ["Project No.", "Document No.", "Revision", "Event_Key", "Source File", "Source Sheet", "Source Row", "Original_Hyperlink_Target", "Source Workbook Path"],
        "documents": ["Project No.", "Document No.", "Global Document Key", "Source File", "Source Sheet", "Source Row", "Source Workbook Path", "Original_Hyperlink_Target"],
        "revisions": ["Project No.", "Document No.", "Revision", "Revision Key", "Source File", "Source Sheet", "Source Row", "Source Workbook Path", "Original_Hyperlink_Target"]}
class WorkspaceLock:
    def __init__(self, root):
        self.f = (root / ".writer.lock").open("a+b")
        self.f.seek(0)
        if not self.f.read(1):
            self.f.write(b"0"); self.f.flush()
        self.f.seek(0)
        try:
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(self.f.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            self.f.close()
            raise ValueError("This workspace is already open.")
    def close(self): self.f.close()

def read_json(p, default):
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else default

def ordered_fields(rows, scope):
    base = list(OUTPUT_FIELDS) if scope == "events" else []
    return list(dict.fromkeys(base + [k for r in rows for k in r] + CORE[scope]))

def filter_rows(rows, query="", column="", sort="", desc=False):
    tokens = shlex.split(query.replace("+", " ")) if query else []
    def matches(row):
        text = str(row.get(column, "")).casefold() if column else " ".join(str(v) for v in row.values()).casefold()
        return all((t[1:].casefold() not in text) if t.startswith("-") else (t.casefold() in text) for t in tokens)
    result = [r for r in rows if matches(r)] if tokens else list(rows)
    if sort: result.sort(key=lambda r: str(r.get(sort, "")).casefold(), reverse=desc)
    return result

def export_bytes(rows, fields, fmt):
    if fmt == "json":
        return json.dumps([{k:r.get(k,"") for k in fields} for r in rows], ensure_ascii=False, default=str).encode("utf-8")
    if fmt == "csv":
        b=io.StringIO(newline=""); w=csv.DictWriter(b,fieldnames=fields); w.writeheader()
        # CSV has no text-cell metadata: neutralize spreadsheet-formula prefixes explicitly.
        def safe(v):
            s=str(v) if v is not None else ""
            return "'" + s if s.lstrip().startswith(("=", "+", "-", "@")) else s
        w.writerows({k:safe(r.get(k,"")) for k in fields} for r in rows)
        return b.getvalue().encode("utf-8-sig")
    if fmt != "xlsx": raise ValueError("Unknown export format")
    if len(rows)>1048575 or len(fields)>16384: raise ValueError("Excel limits exceeded; partitioning is not enabled in this candidate. Use JSON or CSV.")
    if not rows: raise ValueError("No records to export in this candidate.")
    if "Document Link" in fields and sum(isinstance(r.get("Document Link"),str) and r["Document Link"].startswith(("https://","http://")) for r in rows)>65530:
        raise ValueError("Native hyperlink limit exceeded; no links will be silently omitted. Use JSON/CSV until partitioning is implemented.")
    out=io.BytesIO()
    with xlsxwriter.Workbook(out, {"in_memory":True,"strings_to_formulas":False,"strings_to_urls":False}) as book:
        ws=book.add_worksheet("INDEX"); date_fmt=book.add_format({"num_format":"dd/mm/yyyy"}); links=0
        for i,row in enumerate(rows,1):
            for j,k in enumerate(fields):
                v=row.get(k,"")
                if isinstance(v,(dict,list)): v=json.dumps(v,ensure_ascii=False)
                if isinstance(v,str) and len(v)>32767: raise ValueError("Cell exceeds Excel text limit: "+k)
                if k in ("Event Date","Latest Event Date","Source Modified Date") and isinstance(v,str) and len(v)==10:
                    try: dt=datetime.strptime(v,"%Y-%m-%d")
                    except ValueError: pass
                    else: ws.write_datetime(i,j,dt,date_fmt); continue
                if k=="Document Link" and isinstance(v,str) and v.startswith(("https://","http://")) and links<65530 and len(v)<=2079:
                    if ws.write_url(i,j,v)==0: links+=1; continue
                # All other paths remain literal until real Excel link behavior is verified.
                rc=ws.write(i,j,v)
                if rc != 0: raise ValueError("Excel cell write rejected: "+k)
        if ws.add_table(0,0,len(rows),len(fields)-1,{"name":"NMDCIndex","columns":[{"header":k} for k in fields]}) != 0:
            raise ValueError("Excel table creation failed")
        ws.freeze_panes(1,0); ws.set_column(0,len(fields)-1,18)
    return out.getvalue()

class App:
    def __init__(self, workspace):
        self.root=workspace.resolve(); self.root.mkdir(parents=True,exist_ok=True)
        self.lock=WorkspaceLock(self.root)
        self.state=self.root/"state"; self.config=self.root/"config"; self.config.mkdir(exist_ok=True)
        for p in (ROOT/"config").glob("*.csv"):
            dest=self.config/p.name
            if not dest.exists(): shutil.copy2(p,dest)
        self.settings_path=self.root/"settings.json"
        self.settings=read_json(self.settings_path, {})
        self.guard=threading.RLock(); self.job={"status":"idle"}; self.cache={}; self.token=secrets.token_urlsafe(32)
    def save_settings(self): _write_json_atomic(self.settings_path,self.settings)
    def source(self):
        v=self.settings.get("data_dir")
        if not v: raise ValueError("Select the DATA folder first.")
        p=Path(v).resolve()
        if not p.is_dir(): raise ValueError("Source folder is unavailable.")
        if self.root==p or self.root in p.parents or p in self.root.parents:
            raise ValueError("Source and workspace must be separate.")
        return p
    def idle(self):
        if self.job["status"]=="running": raise ValueError("A processing job is active.")
    def select(self):
        self.idle()
        import tkinter as tk
        from tkinter import filedialog
        root=tk.Tk(); root.withdraw(); root.attributes("-topmost",True)
        try: value=filedialog.askdirectory(parent=root,title="Select NMDC DATA folder",mustexist=True)
        finally: root.destroy()
        if not value: return self.settings
        previous=self.settings.get("data_dir")
        if previous and Path(previous).resolve()!=Path(value).resolve() and (self.state/"approved"/"current.json").exists():
            raise ValueError("Existing approved data belongs to another source. Use an isolated workspace.")
        self.settings["data_dir"]=value
        try:self.source()
        except Exception:
            if previous: self.settings["data_dir"]=previous
            else:self.settings.pop("data_dir",None)
            raise
        self.save_settings(); return self.settings
    def start(self,full):
        self.idle(); src=self.source()
        self.job={"status":"running","mode":"full" if full else "incremental","started":time.time()}
        def work():
            try:
                install_resilient_source_access(); install_layout_compatibility()
                result=stage_runtime_update(data_dir=src,state_dir=self.state,config_dir=self.config,full_rescan=full)
                with self.guard:self.job={"status":"complete","result":result}
            except Exception as exc:
                with self.guard:self.job={"status":"error","message":str(exc)}
        threading.Thread(target=work,daemon=True).start()
        return self.job
    def approved(self):
        marker=self.state/"approved"/"current.json"
        key=marker.read_bytes() if marker.exists() else b"legacy"
        if self.cache.get("key")!=key:
            a=load_approved_state(self.state)
            self.cache={"key":key,"approved":a,"views":{}}
        return self.cache["approved"]
    def view(self,scope):
        a=self.approved(); views=self.cache["views"]
        if scope not in CORE: raise ValueError("Unknown view")
        if scope not in views:
            raw=a["records"]
            rows=[dict(r) for r in (raw if scope=="events" else build_document_rows(raw) if scope=="documents" else build_revision_rows(raw))]
            source_root=a.get("manifest",{}).get("data_dir","")
            if scope!="events":
                by_source={(r.get("Source File"),r.get("Source Sheet"),str(r.get("Source Row"))):r for r in raw}
                for r in rows:
                    orig=by_source.get((r.get("Source File"),r.get("Source Sheet"),str(r.get("Source Row"))),{})
                    r["Original_Hyperlink_Target"]=orig.get("Original_Hyperlink_Target","")
            for r in rows:
                r["Source Workbook Path"]=str(Path(source_root)/str(r.get("Source File",""))) if source_root else str(r.get("Source File",""))
            views[scope]=rows
        return views[scope]
    def proposal(self):
        self.idle(); s=load_latest_stage(self.state)
        run=s.get("run_id",""); status=s.get("decision",{}).get("decision","")
        if not run or status in ("APPROVED","REJECTED"): return {"run":"","files":[],"flags":[]}
        files=affected_files(self.state,run)
        saved=read_json(s["stage_dir"]/"file_review.json",{})
        valid=False
        if saved:
            try: validated_review(self.state,run,required=True); valid=True
            except ValueError: pass
        if valid:
            for f in files:
                c=saved["choices"][f["Source File"]]; f["Accept Update?"]=c["accept"]; f["Owner Note"]=c["note"]
        return {"run":run,"files":files,"flags":build_flag_rows(s),"summary":s["summary"],"saved":valid}
    def save_choices(self,body):
        self.idle()
        p=self.proposal()
        if body.get("run")!=p["run"] or not p["run"]: raise ValueError("Proposal has changed")
        f=self.root/"choices.csv"
        try:
            with f.open("w",encoding="utf-8",newline="") as h:
                w=csv.DictWriter(h,fieldnames=["Source File","Accept Update?","Owner Note"]); w.writeheader()
                for r in body["files"]:
                    if type(r["Accept Update?"]) is not bool: raise ValueError("Invalid choice")
                    w.writerow({"Source File":r["Source File"],"Accept Update?":str(r["Accept Update?"]).upper(),"Owner Note":str(r.get("Owner Note",""))})
            return save_file_review(self.state,p["run"],f)
        finally: f.unlink(missing_ok=True)
    def approve(self,body):
        self.idle(); p=self.proposal()
        if not p["run"] or body.get("run")!=p["run"]: raise ValueError("Proposal changed")
        # Conservative candidate: no override and no unresolved non-routine flag approval.
        routine={"SOURCE_NEW","SOURCE_CHANGED","PARSER_VERSION_CHANGED","CONFIG_CHANGED","CACHE_MISSING"}
        if any(f["Flag Code"] not in routine and f.get("Resolution Status")!="RESOLVED" for f in p["flags"]):
            raise ValueError("Unresolved flags require correction; candidate does not provide override.")
        result=approve_stage(self.state,p["run"],require_file_review=True,allow_conflicts=False)
        self.cache={};return result

def server_for(app):
    class Handler(BaseHTTPRequestHandler):
        server_version="NMDC-V2"
        def log_message(self,*args): pass
        def send(self,code,payload,ctype="application/json",name=None):
            if not isinstance(payload,bytes): payload=json.dumps(payload,ensure_ascii=False,default=str).encode()
            self.send_response(code); self.send_header("Content-Type",ctype)
            self.send_header("Content-Length",str(len(payload))); self.send_header("Cache-Control","no-store")
            self.send_header("X-Content-Type-Options","nosniff"); self.send_header("Referrer-Policy","no-referrer")
            self.send_header("Content-Security-Policy","default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'")
            if name:self.send_header("Content-Disposition",'attachment; filename="'+name+'"')
            self.end_headers();self.wfile.write(payload)
        def auth(self,api=True):
            expected="127.0.0.1:"+str(self.server.server_port)
            if self.headers.get("Host")!=expected: raise PermissionError("Host rejected")
            origin=self.headers.get("Origin")
            if origin and origin!="http://"+expected: raise PermissionError("Origin rejected")
            if self.command=="POST" and origin!="http://"+expected: raise PermissionError("Origin required")
            if api and not hmac.compare_digest(self.headers.get("Authorization",""),"Bearer "+app.token): raise PermissionError("Invalid session")
        def do_GET(self):
            try:
                u=urlsplit(self.path); path=u.path
                if path in ("/","/app.js","/style.css"):
                    self.auth(False); f=WEB/{" /":"index.html","/":"index.html","/app.js":"app.js","/style.css":"style.css"}[path]
                    return self.send(200,f.read_bytes(),{"/":"text/html; charset=utf-8","/app.js":"application/javascript; charset=utf-8","/style.css":"text/css; charset=utf-8"}[path])
                self.auth()
                q={k:v[0] for k,v in parse_qs(u.query).items()}
                with app.guard:
                    if path=="/api/status":
                        a=app.approved()
                        return self.send(200,{"job":app.job,"settings":app.settings,"approved_run":a["run_id"],"events":len(a["records"])})
                    if path=="/api/review":return self.send(200,app.proposal())
                    if path in ("/api/index","/api/export"):
                        scope=q.get("scope","events"); rows=app.view(scope); fields=ordered_fields(rows,scope)
                        rows=filter_rows(rows,q.get("q",""),q.get("column",""),q.get("sort",""),q.get("desc")=="1")
                        if path=="/api/index":
                            page=max(0,int(q.get("page","0")));return self.send(200,{"count":len(rows),"fields":fields,"rows":rows[page*100:(page+1)*100]})
                        selected=q.get("columns","").split("|") if q.get("columns") else fields
                        selected=list(dict.fromkeys(selected+CORE[scope]))
                        if not set(selected)<=set(fields): raise ValueError("Unknown column")
                        fmt=q.get("format","xlsx"); data=export_bytes(rows,selected,fmt)
                        return self.send(200,data,{"xlsx":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","csv":"text/csv","json":"application/json"}[fmt],"NMDC_DOCUMENT_INDEX."+fmt)
                    return self.send(404,{"error":"Unknown route"})
            except PermissionError as e:self.send(403,{"error":str(e)})
            except (ValueError,OSError,KeyError) as e:self.send(400,{"error":str(e)})
            except Exception as e:self.send(500,{"error":str(e)})
        def do_POST(self):
            try:
                self.auth()
                size=int(self.headers.get("Content-Length","0"))
                if not 0<=size<=2_000_000: raise ValueError("Request too large")
                if not self.headers.get("Content-Type","").startswith("application/json"):raise ValueError("JSON required")
                body=json.loads(self.rfile.read(size) or b"{}")
                with app.guard:
                    path=urlsplit(self.path).path
                    if path=="/api/select":result=app.select()
                    elif path=="/api/scan":result=app.start(body.get("full") is True)
                    elif path=="/api/save-choices":result=app.save_choices(body)
                    elif path=="/api/approve":result=app.approve(body)
                    elif path=="/api/undo":
                        app.idle();result=undo_last_approval(app.state);app.cache={}
                    elif path=="/api/preferences":
                        if body.get("theme") not in ("light","dark"):raise ValueError("Invalid theme")
                        app.settings["theme"]=body["theme"];app.save_settings();result=app.settings
                    elif path=="/api/exit":
                        app.idle();result={"status":"closing"};threading.Thread(target=self.server.shutdown,daemon=True).start()
                    else:return self.send(404,{"error":"Unknown route"})
                    return self.send(200,result)
            except PermissionError as e:self.send(403,{"error":str(e)})
            except (ValueError,OSError,KeyError) as e:self.send(400,{"error":str(e)})
            except Exception as e:self.send(500,{"error":str(e)})
    s=ThreadingHTTPServer(("127.0.0.1",0),Handler)
    return s

def self_test():
    import tempfile, zipfile, urllib.request
    with tempfile.TemporaryDirectory() as d:
        app=App(Path(d)/"workspace")
        s=server_for(app);t=threading.Thread(target=s.serve_forever,daemon=True);t.start()
        try:
            url="http://127.0.0.1:"+str(s.server_port)
            req=urllib.request.Request(url+"/api/status",headers={"Authorization":"Bearer "+app.token})
            assert json.load(urllib.request.urlopen(req))["events"]==0
            assert b"NMDC" in urllib.request.urlopen(url).read()
            try: urllib.request.urlopen(url+"/api/status")
            except urllib.error.HTTPError as e: assert e.code==403
            else:raise AssertionError("Missing token accepted")
            rows=[{"Event_Key":"E1","Document No.":"0001","Document Title":"=1+1","Document Link":"https://example.invalid/a","Event Date":"2026-11-28","Source File":"test.xlsx"}]
            b=export_bytes(rows,list(rows[0]),"xlsx")
            with zipfile.ZipFile(io.BytesIO(b)) as z:
                assert z.testzip() is None and "xl/tables/table1.xml" in z.namelist()
                assert b'NMDCIndex' in z.read("xl/tables/table1.xml")
                assert b'<f>' not in z.read("xl/worksheets/sheet1.xml")
            assert filter_rows(rows,"0001 -missing","Document No.")==rows
            # Exercise the pinned production update and review engine with an isolated synthetic processor.
            data=Path(d)/"DATA";data.mkdir();source=data/"register.xlsx";source.write_bytes(b"synthetic input")
            before=hashlib.sha256(source.read_bytes()).hexdigest()
            def processor(path,rel):
                return [{"Event_Key":"E1","Source File":rel,"Project No.":"0001","Document No.":"D1","Document Title":"Fixture","Revision":"A","Revision_Key":"R1","Global_Document_Key":"D1","Source Sheet":"Register","Source Row":2,"Source Cell":"A2","Event Date":"2026-11-28"}]
            staged=stage_update(data_dir=data,state_dir=app.state,processor=processor)
            p=app.proposal();assert p["run"]==staged["run_id"] and len(p["files"])==1
            try:app.approve({"run":p["run"]})
            except ValueError:pass
            else:raise AssertionError("Unsaved approval permitted")
            app.save_choices({"run":p["run"],"files":p["files"]})
            app.approve({"run":p["run"]})
            assert app.approved()["records"][0]["Event_Key"]=="E1"
            assert hashlib.sha256(source.read_bytes()).hexdigest()==before
            assert app.view("events")[0]["Document No."]=="D1"
            print("Standalone HTTP/auth/XLSX and production update/review smoke PASS")
        finally:s.shutdown();s.server_close();app.lock.close()
def main():
    p=argparse.ArgumentParser();p.add_argument("--self-test",action="store_true");p.add_argument("--workspace",type=Path)
    args=p.parse_args()
    if args.self_test:self_test();return
    root=args.workspace or Path(os.environ.get("LOCALAPPDATA",Path.home()))/"NMDCIndexV2"
    app=App(root)
    try:
        s=server_for(app);url="http://127.0.0.1:"+str(s.server_port)+"/#"+app.token
        print("NMDC V2 running. Close using Exit in the app. "+url,flush=True);webbrowser.open(url)
        s.serve_forever()
    finally:
        if "s" in locals():s.server_close()
        app.lock.close()
if __name__=="__main__":main()
