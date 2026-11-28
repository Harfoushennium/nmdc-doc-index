"""Repeated packaged HTTP operations; read-only corpus, disposable workspaces.

Cold means fresh process/workspace, not an OS cache flush. The five warm trials
reuse that workspace. Report limits without silently waiving a missed ceiling.
"""
import argparse
import hashlib
import json
import os
import platform
import statistics
import subprocess
import tempfile
import time
import urllib.request
from pathlib import Path
from urllib.parse import urlencode


def source_hashes(data):
    return {p.relative_to(data).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in data.rglob('*') if p.is_file() and p.suffix.lower() in ('.xlsx', '.xlsm') and not p.name.startswith('~$')}


def peak_memory(pid):
    if os.name=="nt":
        import ctypes
        from ctypes import wintypes
        class Counters(ctypes.Structure):
            _fields_=[("cb",wintypes.DWORD),("faults",wintypes.DWORD)]+[(name,ctypes.c_size_t) for name in ("peak_working_set","working_set","peak_paged","paged","peak_nonpaged","nonpaged","pagefile","peak_pagefile")]
        kernel=ctypes.WinDLL("kernel32",use_last_error=True);psapi=ctypes.WinDLL("psapi",use_last_error=True)
        kernel.OpenProcess.restype=wintypes.HANDLE;kernel.CloseHandle.argtypes=[wintypes.HANDLE]
        psapi.GetProcessMemoryInfo.argtypes=[wintypes.HANDLE,ctypes.POINTER(Counters),wintypes.DWORD]
        handle=kernel.OpenProcess(0x410,False,pid)
        if not handle:return None
        try:
            counters=Counters();counters.cb=ctypes.sizeof(counters)
            if not psapi.GetProcessMemoryInfo(handle,ctypes.byref(counters),counters.cb):return None
            return counters.peak_working_set
        finally:kernel.CloseHandle(handle)
    status_path=Path(f"/proc/{int(pid)}/status")
    try:
        for line in status_path.read_text(encoding="utf-8",errors="ignore").splitlines():
            if line.startswith("VmHWM:"):
                parts=line.split()
                if len(parts)>=2 and parts[1].isdigit():
                    unit=parts[2].lower() if len(parts)>=3 else "kb"
                    mult={"b":1,"kb":1024,"mb":1024**2,"gb":1024**3}.get(unit,1024)
                    return int(parts[1])*mult
    except OSError:
        return None
    return None


def run(executable, data, output):
    before=source_hashes(data)
    assert len(before)==56, 'Unexpected corpus; verify its identity before benchmarking'
    report={'method':'3 fresh-process/workspace trials + 5 same-process/workspace warm trials; no OS cache flush',
            'platform':platform.platform(),'executable_sha256':hashlib.sha256(executable.read_bytes()).hexdigest(),
            'workbooks':len(before),'cold':[],'warm':[],'queries':[],'approved_restart_seconds':[]}
    def measure(operation):
        start=time.perf_counter();result=operation()
        return result,round(time.perf_counter()-start,6)
    with tempfile.TemporaryDirectory(prefix='nmdc-release-benchmark-') as directory:
        root=Path(directory)
        for cold in range(3):
            workspace=root/str(cold);workspace.mkdir()
            (workspace/'settings.json').write_text(json.dumps({'data_dir':str(data)}),encoding='utf-8')
            env=os.environ.copy();env.pop('PYTHONPATH',None)
            system=env.get('SystemRoot',r'C:\Windows');env['PATH']=os.pathsep.join([system,str(Path(system)/'System32')])
            started=time.perf_counter()
            process=subprocess.Popen([str(executable),'--workspace',str(workspace),'--no-browser'],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,env=env)
            try:
                line=process.stdout.readline();assert 'http://127.0.0.1:' in line,'Packaged startup failed'
                url,token=line[line.index('http://'):].strip().split('#',1);url=url.rstrip('/')
                def api(path,body=None,raw=False):
                    headers={'Authorization':'Bearer '+token}
                    if body is not None:headers.update({'Origin':url,'Content-Type':'application/json'})
                    req=urllib.request.Request(url+path,headers=headers,data=None if body is None else json.dumps(body).encode())
                    with urllib.request.urlopen(req,timeout=180) as response:
                        return response.read() if raw else json.load(response)
                api('/api/status');startup=time.perf_counter()-started
                def scan(full):
                    api('/api/scan',{'full':full});deadline=time.monotonic()+180
                    while time.monotonic()<deadline:
                        job=api('/api/status')['job']
                        if job['status']!='running':
                            assert job['status']=='complete',job
                            return job['result']
                        time.sleep(.025)
                    raise TimeoutError('Packaged scan timed out')
                for trial in range(6 if cold==2 else 1):
                    result,full=measure(lambda:scan(True))
                    assert result['staged_records']==40494 and not result['blocking_flags'],result
                    proposal=api('/api/review');api('/api/save-choices',proposal)
                    proposal=api('/api/review')
                    approval,approve=measure(lambda:api('/api/approve',{k:proposal[k] for k in ('run','draft_version','flag_version')}))
                    assert approval.get('verified_unchanged_eligible'),approval
                    unchanged,noop=measure(lambda:scan(False));assert unchanged['status']=='UNCHANGED'
                    query={'scope':'events','format':'xlsx'}
                    def export():
                        plan=api('/api/export-plan?'+urlencode(query))
                        args=dict(query,preview=plan['preview'],run=plan['run'],literal_links='confirmed',literal_dates='confirmed',partition='confirmed')
                        content=api('/api/export?'+urlencode(args),raw=True)
                        target=workspace/'measured-export.zip';target.write_bytes(content)
                        assert content[:2]==b'PK';return len(content)
                    size,export_seconds=measure(export)
                    row={'startup':round(startup,6) if trial==0 else None,'full':full,'unchanged':noop,
                         'approval':approve,'export':export_seconds,'export_bytes':size,'events':40494,'parent_peak_rss_bytes':peak_memory(process.pid)}
                    report['cold' if trial==0 else 'warm'].append(row)
                    print(json.dumps({'process':cold,'trial':trial,**row}),flush=True)
                    if cold==2 and trial==5:
                        for index in range(45):
                            scope=('events','documents','revisions')[index%3]
                            result,elapsed=measure(lambda:api('/api/index?'+urlencode({'scope':scope,'q':('-absentfixture' if index%2 else '')})))
                            report['queries'].append(elapsed)
                            assert result['count']>0
                api('/api/exit',{});assert process.wait(timeout=15)==0
                restart_start=time.perf_counter()
                process=subprocess.Popen([str(executable),'--workspace',str(workspace),'--no-browser'],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,env=env)
                line=process.stdout.readline();assert 'http://127.0.0.1:' in line,'Approved-workspace restart failed'
                url,token=line[line.index('http://'):].strip().split('#',1);url=url.rstrip('/')
                assert api('/api/status')['events']==40494
                assert api('/api/index?scope=events')['count']==40494
                report['approved_restart_seconds'].append(round(time.perf_counter()-restart_start,6))
                api('/api/exit',{});assert process.wait(timeout=15)==0
            finally:
                if process.poll() is None:process.terminate();process.wait(timeout=15)
        report['scratch_disk_bytes']=sum(p.stat().st_size for p in root.rglob('*') if p.is_file())
    assert before==source_hashes(data),'Source corpus changed'
    report['source_hashes_unchanged']=True
    all_trials=report['cold']+report['warm']
    report['medians']={k:statistics.median(row[k] for row in all_trials) for k in ('full','unchanged','approval','export')}
    report['query_p95']=sorted(report['queries'])[int(.95*len(report['queries']))]
    report['release_ceilings']={
        'startup':max([row['startup'] for row in report['cold']]+report['approved_restart_seconds'])<=5,
        'full':report['medians']['full']<=15,'unchanged':report['medians']['unchanged']<=3,
        'export':report['medians']['export']<=10,'query':report['query_p95']<=.5}
    report['owner_pc_acceptance']='NOT RUN; local development environment only'
    output.write_text(json.dumps(report,indent=2),encoding='utf-8')
    print(json.dumps({k:report[k] for k in ('medians','query_p95','release_ceilings','source_hashes_unchanged')},indent=2))


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('executable',type=Path)
    parser.add_argument('--data',type=Path,default=Path(__file__).resolve().parents[1]/'DATA')
    parser.add_argument('--output',type=Path,default=Path('v2-release-benchmark.json'))
    args=parser.parse_args();run(args.executable.resolve(),args.data.resolve(),args.output)
