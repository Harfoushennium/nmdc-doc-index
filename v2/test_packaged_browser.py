"""Windows CI browser acceptance of the packaged candidate, using a disposable synthetic stage."""
import os, sys, tempfile, subprocess, zipfile
from pathlib import Path
from playwright.sync_api import sync_playwright, expect
from nmdc_profiler.update_engine import stage_update
from v2.launcher import App

def test(exe):
    with tempfile.TemporaryDirectory(prefix="nmdc-ui-") as tmp:
        root=Path(tmp);workspace=root/"workspace";data=root/"DATA";data.mkdir()
        (data/"register.xlsx").write_bytes(b"synthetic")
        app=App(workspace)
        try:
            stage_update(data_dir=data,state_dir=app.state,processor=lambda path,rel:[{
                "Event_Key":"E-UI-1","Global_Document_Key":"D-UI-1","Revision_Key":"R-UI-1",
                "Document No.":"UI-0001","Document Title":"Synthetic browser smoke",
                "Project No.":"0001","Revision":"A","Source File":rel,"Source Sheet":"Register",
                "Source Row":2,"Source Cell":"A2","Event Date":"2026-11-28"}])
        finally:app.lock.close()
        env=os.environ.copy();env.pop("PYTHONPATH",None)
        env["PATH"]=os.pathsep.join([str(Path(env["SystemRoot"])/"System32"),env["SystemRoot"]])
        proc=subprocess.Popen([str(exe),"--workspace",str(workspace),"--no-browser"],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,env=env)
        try:
            line=proc.stdout.readline()
            assert "http://127.0.0.1:" in line, "Standalone service failed to start"
            url=line[line.index("http://"):].strip()
            with sync_playwright() as pw:
                browser=pw.chromium.launch(channel="msedge",headless=True)
                page=browser.new_page(accept_downloads=True)
                page.on("dialog",lambda dialog:dialog.accept())
                page.goto(url)
                expect(page.locator("#title")).to_have_text("Dashboard")
                page.get_by_role("button",name="File review & flags",exact=True).click()
                note=page.get_by_placeholder("Owner note")
                expect(note).to_be_visible();note.fill("Retained across reload")
                expect(page.locator("#notice")).to_contain_text("Draft retained")
                page.reload()
                page.get_by_role("button",name="File review & flags",exact=True).click()
                expect(page.get_by_placeholder("Owner note")).to_have_value("Retained across reload")
                expect(page.get_by_role("button",name="Approve saved choices",exact=True)).to_be_disabled()
                page.get_by_role("button",name="Save file choices",exact=True).click()
                expect(page.get_by_role("button",name="Approve saved choices",exact=True)).to_be_enabled()
                page.get_by_role("button",name="Approve saved choices",exact=True).click()
                expect(page.get_by_text("No pending proposal",exact=True)).to_be_visible()
                page.get_by_role("button",name="Consolidated index",exact=True).click()
                expect(page.get_by_text("UI-0001",exact=True)).to_be_visible()
                page.get_by_role("button",name="Export",exact=True).click()
                with page.expect_download() as download:
                    page.get_by_role("button",name="Download XLSX",exact=True).click()
                path=root/"export.xlsx";download.value.save_as(path)
                with zipfile.ZipFile(path) as z:
                    assert z.testzip() is None and "xl/tables/table1.xml" in z.namelist()
                page.get_by_role("button",name="Exit application",exact=True).click()
                expect(page.get_by_text("Application closed.",exact=True)).to_be_visible()
                browser.close()
            assert proc.wait(timeout=15)==0
            print("Packaged browser interaction PASS: draft/reload/save/approve/index/export/exit")
        finally:
            if proc.poll() is None:proc.terminate();proc.wait(timeout=10)
if __name__=="__main__":test(Path(sys.argv[1]).resolve())
