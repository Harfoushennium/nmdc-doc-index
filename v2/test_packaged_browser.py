"""Windows CI browser acceptance of the packaged candidate, using a disposable synthetic stage."""

import os, sys, tempfile, subprocess, zipfile

from pathlib import Path

from playwright.sync_api import sync_playwright, expect

from nmdc_profiler.update_engine import stage_update

from v2.launcher import App



def test(exe):

    with tempfile.TemporaryDirectory(prefix="nmdc-ui-") as tmp, tempfile.TemporaryDirectory(prefix="nmdc-reports-") as reports:

        root=Path(tmp);workspace=root/"workspace";data=root/"DATA";data.mkdir()

        (data/"register.xlsx").write_bytes(b"synthetic")

        app=App(workspace)

        try:

            app.settings["report_parent"]=reports;app.save_settings()

            stage_update(data_dir=data,state_dir=app.state,processor=lambda path,rel:[{

                "Event_Key":"E-UI-1","Global_Document_Key":"D-UI-1","Revision_Key":"R-UI-1",

                "Document No.":"UI-0001","Document Title":"Synthetic browser smoke",

                "Project No.":"0001","Revision":"A","Source File":rel,"Source Sheet":"Register",

                "Source Row":2,"Source Cell":"A2","Event Date":"2026-11-28"}],initial_flags=[{"level":"REVIEW","code":"SOURCE_NEW","source":"register.xlsx","source_sheet":"Register","source_row":2,"source_cell":"A2","message":"Fixture source review"}])

        finally:app.lock.close()

        env=os.environ.copy();env.pop("PYTHONPATH",None)

        system_root=env.get("SYSTEMROOT",env.get("SystemRoot",env.get("WINDIR",r"C:\Windows")))

        env["PATH"]=os.pathsep.join([str(Path(system_root)/"System32"),system_root])

        proc=subprocess.Popen([str(exe),"--workspace",str(workspace),"--no-browser"],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,env=env)

        try:

            line=proc.stdout.readline()

            assert "http://127.0.0.1:" in line, "Standalone service failed to start"

            url=line[line.index("http://"):].strip()

            with sync_playwright() as pw:

                browser=pw.chromium.launch(channel="msedge",headless=True)

                page=browser.new_page(accept_downloads=True)

                page.on("dialog",lambda dialog:dialog.accept("Browser saved" if dialog.type=="prompt" else ""))

                page.goto(url)

                expect(page.locator("#title")).to_have_text("Dashboard")

                page.get_by_role("button",name="File review & flags",exact=True).click()

                page.get_by_text("Inspect record differences",exact=True).click()

                expect(page.get_by_text("1 changed records",exact=True)).to_be_visible()

                page.get_by_text("ADDED · EVENT:E-UI-1",exact=True).click()

                expect(page.get_by_text("Register · Row 2",exact=False)).to_be_visible()

                note=page.get_by_placeholder("Owner note")

                expect(note).to_be_visible();note.fill("Retained across reload")

                expect(page.locator("#notice")).to_contain_text("Draft retained")

                page.reload()

                page.get_by_role("button",name="File review & flags",exact=True).click()

                expect(page.get_by_placeholder("Owner note")).to_have_value("Retained across reload")

                expect(page.get_by_role("button",name="Approve saved choices",exact=True)).to_be_disabled()

                page.get_by_role("button",name="Save file choices",exact=True).click()

                expect(page.get_by_role("button",name="Approve saved choices",exact=True)).to_be_enabled()

                flag_note=page.get_by_placeholder("Flag owner comment").first

                flag_note.fill("Checked source evidence")

                expect(page.locator("#notice")).to_contain_text("Flag draft retained")

                page.reload()

                page.get_by_role("button",name="File review & flags",exact=True).click()

                expect(page.get_by_placeholder("Flag owner comment").first).to_have_value("Checked source evidence")

                expect(page.get_by_role("button",name="Approve saved choices",exact=True)).to_be_disabled()

                page.get_by_role("button",name="Save flag decisions",exact=True).click()

                expect(page.get_by_role("button",name="Approve saved choices",exact=True)).to_be_enabled()

                page.locator("label").filter(has_text="SOURCE_NEW").locator("input[type=checkbox]").first.check()

                page.get_by_role("button",name="Generate parser/mapping fix report",exact=True).click()

                expect(page.locator("#notice")).to_contain_text("Report created:")

                assert list(Path(reports).glob("PARSER_FIX_REPORTS/*/PARSER_FIX_REQUEST.json"))

                page.get_by_role("button",name="Clear decision and comment",exact=True).first.click()

                page.get_by_role("button",name="Save flag decisions",exact=True).click()

                expect(page.get_by_role("button",name="Approve saved choices",exact=True)).to_be_enabled()

                page.get_by_role("button",name="Approve saved choices",exact=True).click()

                expect(page.get_by_text("No pending proposal",exact=True)).to_be_visible()

                page.get_by_role("button",name="Consolidated index",exact=True).click()

                expect(page.get_by_text("UI-0001",exact=True)).to_be_visible()

                search=page.get_by_placeholder("Search +AND / -EXCLUDE")

                search.fill("no-match")

                expect(page.get_by_text("0 matching records",exact=True)).to_be_visible()

                search.fill('"UI-0001" -absent')

                expect(page.get_by_text("1 matching records",exact=True)).to_be_visible()

                expect(search).to_be_focused()

                page.get_by_role("button",name="Reset filters",exact=True).click()

                expect(page.get_by_placeholder("Search +AND / -EXCLUDE")).to_have_value("")

                page.get_by_role("button",name="Dashboard",exact=True).click()

                expect(page.get_by_text("Approved documents",exact=True)).to_be_visible()

                expect(page.get_by_text("Documents by Source Family",exact=True)).to_be_visible()

                page.get_by_role("button",name="Rules & mappings",exact=True).click()

                rules=page.locator("table").first

                old_note=rules.locator("tbody tr").first.locator("input").nth(17).input_value()

                rules.locator("tbody tr").first.locator("input").nth(17).fill("Browser saved rule")

                page.get_by_role("button",name="Save rules",exact=True).click()

                expect(page.locator("#notice")).to_contain_text("Rules saved")

                page.reload()

                page.get_by_role("button",name="Rules & mappings",exact=True).click()

                expect(page.locator("table").first.locator("tbody tr").first.locator("input").nth(17)).to_have_value("Browser saved rule")

                page.get_by_role("button",name="Restore this configuration",exact=True).first.click()

                expect(page.locator("table").first.locator("tbody tr").first.locator("input").nth(17)).to_have_value(old_note)

                cf=page.locator(".card").filter(has=page.get_by_role("heading",name="Custom Fields & Keywords — derived Document view",exact=True))

                cf.get_by_role("button",name="Add row",exact=True).click()

                cf.locator("tbody tr").first.locator("td").nth(1).locator("input").fill("Browser field")

                km=page.locator(".card").filter(has=page.get_by_role("heading",name="Keyword mappings",exact=True))

                km.get_by_role("button",name="Add row",exact=True).click()

                mapping=km.locator("tbody tr").first

                mapping.locator("td").nth(1).locator("input").fill("Browser field")

                mapping.locator("td").nth(2).locator("input").fill("Synthetic")

                mapping.locator("td").nth(3).locator("input").fill("Browser result")

                page.get_by_role("button",name="Save and apply custom fields",exact=True).click()

                expect(page.locator("#notice")).to_contain_text("Derived fields saved")

                page.get_by_role("button",name="Consolidated index",exact=True).click()

                page.locator("#content select").first.select_option("documents")

                expect(page.get_by_role("cell",name="Browser result",exact=True)).to_be_visible()

                page.get_by_role("button",name="Save index view",exact=True).click()

                expect(page.locator("#notice")).to_contain_text("Saved index view")

                page.get_by_role("button",name="Settings",exact=True).click()

                appearance=page.locator(".card").filter(has=page.get_by_role("heading",name="Appearance",exact=True))

                appearance.locator("select").first.select_option("dark")

                expect(page.locator("body")).to_have_class("dark")

                page.reload()

                expect(page.locator("body")).to_have_class("dark")

                page.locator("nav").get_by_role("button",name="Export",exact=True).click()

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

