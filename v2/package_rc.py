"""Assemble and verify a portable candidate; never publish Production/Latest."""
import argparse
import csv
import hashlib
import importlib.metadata
import json
import platform
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def assemble(built, destination):
    root=Path(__file__).resolve().parents[1]
    if destination.exists():raise ValueError('Candidate destination already exists')
    if not (built/'Launch.exe').is_file():raise ValueError('Build Launch.exe first')
    commit=subprocess.check_output(['git','rev-parse','HEAD'],cwd=root,text=True).strip()
    dirty=subprocess.check_output(['git','status','--porcelain','--untracked-files=no'],cwd=root,text=True)
    if dirty.strip():raise ValueError('Commit tracked release changes before pinning the candidate')
    shutil.copytree(built,destination)
    for name in ('Run-UAT.ps1','UAT-GUIDE.md','README.md'):
        shutil.copy2(root/'v2'/name,destination/name)
    notices=destination/'licenses';notices.mkdir()
    shutil.copy2(Path(sys.base_prefix)/'LICENSE.txt',notices/'Python-LICENSE.txt')
    for package in ('xlsxwriter','pyinstaller'):
        distribution=importlib.metadata.distribution(package)
        found=False
        for item in distribution.files or []:
            if any(term in item.name.lower() for term in ('license','copying','notice')) and item.suffix.lower() not in ('.py','.pyc'):
                path=distribution.locate_file(item)
                if path.is_file():shutil.copy2(path,notices/(package+'-'+item.name));found=True
        if not found:raise ValueError('Missing third-party license: '+package)
    (notices/'NOTICE.txt').write_text('Includes CPython and its bundled third-party components, XlsxWriter and PyInstaller bootloader. See complete license texts. PyInstaller GPL includes its bootloader exception. Playwright is a build/test dependency and is not distributed. Application ownership and release approval remain with the repository owner.\n',encoding='utf-8')
    plan=(root/'docs/architecture/NMDC-V2-FINAL-RELEASE-PLAN.md').read_text(encoding='utf-8')
    cases=re.findall(r'^\| ([FMEBPU]\d{2}) \| (.*?) \| (.*?) \|$',plan,re.M)
    if len(cases)!=46:raise ValueError('Acceptance contract must contain exactly 46 cases')
    with (destination/'ACCEPTANCE.csv').open('w',encoding='utf-8-sig',newline='') as handle:
        writer=csv.writer(handle);writer.writerow(['ID','Verification','Expected','Build','Machine','Status','Actual','Evidence','Finding'])
        writer.writerows([case[0],case[1],case[2],commit,'','NOT RUN','','',''] for case in cases)
    files=[]
    for path in sorted(destination.rglob('*')):
        if path.is_symlink():raise ValueError('Package contains a symlink')
        if not path.is_file():continue
        rel=path.relative_to(destination).as_posix()
        if any(part.casefold() in ('data','state','backups','migration-evidence') for part in path.relative_to(destination).parts):raise ValueError('Operational area in candidate')
        if path.suffix.lower() in ('.xlsx','.xlsm','.jsonl'):raise ValueError('Operational record/workbook in candidate')
        files.append({'path':rel,'bytes':path.stat().st_size,'sha256':digest(path)})
    manifest={'schema':1,'version':'RC1','status':'candidate - owner acceptance pending','commit':commit,
              'toolchain':{'python':platform.python_version(),'pyinstaller':importlib.metadata.version('pyinstaller'),'xlsxwriter':importlib.metadata.version('xlsxwriter')},'files':files}
    (destination/'manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')
    archive=destination.with_suffix('.zip')
    if archive.exists():raise ValueError('Candidate ZIP already exists')
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as output:
        for path in sorted(destination.rglob('*')):
            if path.is_file():output.write(path,path.relative_to(destination).as_posix())
    with zipfile.ZipFile(archive) as check:
        assert check.testzip() is None
        assert set(check.namelist())=={f['path'] for f in files}|{'manifest.json'}
        for entry in files:
            content=check.read(entry['path'])
            assert len(content)==entry['bytes'] and hashlib.sha256(content).hexdigest()==entry['sha256']
    checksum=archive.with_suffix('.zip.sha256');checksum.write_text(digest(archive)+'  '+archive.name+'\n',encoding='ascii')
    print(json.dumps({'archive':str(archive),'sha256':digest(archive),'commit':commit,'files':len(files),'bytes':archive.stat().st_size}))


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('built',type=Path);parser.add_argument('destination',type=Path)
    args=parser.parse_args();assemble(args.built.resolve(),args.destination.resolve())
