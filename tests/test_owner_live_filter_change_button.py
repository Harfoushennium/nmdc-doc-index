import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
LIVE_MODULE = ROOT / "excel" / "vba" / "modNMDC_LiveFilter.bas"


class OwnerLiveFilterChangeButtonTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.live = LIVE_MODULE.read_text(encoding="utf-8-sig")

    def test_change_control_layout_and_idempotent_creation(self):
        live = self.live
        setup = live.split(
            "Private Sub NMDC_EnsureReferenceLiveFilterControls", 1
        )[1].split("End Sub", 1)[0]
        search_at = setup.index('Set area = ws.Range("B3:H3")')
        choose_at = setup.index('Set area = ws.Range("I3:J3")')
        reset_at = setup.index('Set area = ws.Range("K3:L3")')
        self.assertLess(search_at, choose_at)
        self.assertLess(choose_at, reset_at)
        self.assertIn('Private Const CHOOSE_BUTTON_NAME As String = "Btn_Choose_Search"', live)
        self.assertIn("Set chooseButton = ws.Buttons(CHOOSE_BUTTON_NAME)", setup)
        self.assertIn("If chooseButton Is Nothing Then", setup)
        self.assertIn("Set chooseButton = ws.Buttons.Add(", setup)
        self.assertIn("chooseButton.Name = CHOOSE_BUTTON_NAME", setup)
        for geometry in ("Left", "Top", "Width", "Height"):
            self.assertIn(f"chooseButton.{geometry} = area.{geometry}", setup)
        self.assertIn('chooseButton.Caption = "CHANGE COLUMN"', setup)
        self.assertIn("chooseButton.Placement = xlMoveAndSize", setup)

    def test_button_and_shortcut_share_header_chooser_and_keep_cancel_validation(self):
        live = self.live
        self.assertIn('Application.OnKey "^+F", "Ask_User_For_Target_Column"', live)
        self.assertIn(
            'chooseButton.OnAction = "\'" & Replace(ThisWorkbook.Name, "\'", "\'\'") & "\'!NMDC_LiveFilterChooseColumn"',
            live,
        )
        wrapper = live.split("Public Sub NMDC_LiveFilterChooseColumn()", 1)[1].split("End Sub", 1)[0]
        self.assertIn("Ask_User_For_Target_Column", wrapper)
        chooser = live.split("Public Sub Ask_User_For_Target_Column()", 1)[1].split("End Sub", 1)[0]
        self.assertIn('Application.InputBox("Click the HEADER of the column to filter:"', chooser)
        self.assertIn("If UserSelection Is Nothing Then Exit Sub", chooser)
        self.assertIn("Call Setup_New_Target(UserSelection)", chooser)
        target = live.split("Public Sub Setup_New_Target(TheHeader As Range)", 1)[1].split("End Sub", 1)[0]
        self.assertIn("If TheTable Is Nothing Then", target)
        self.assertIn("TheTable.HeaderRowRange", target)
        self.assertIn('Left$(CStr(TheColumn.Name), 7) = "__NMDC_"', target)
        self.assertIn("If TheColumn.DataBodyRange Is Nothing Then", target)

    def test_visible_sheet_activation_still_owns_control_creation(self):
        initialize = self.live.split("Public Sub NMDC_LiveFilterInitialize()", 1)[1].split(
            "Private Sub NMDC_EnsureReferenceLiveFilterControls", 1
        )[0]
        self.assertIn("If Not Application.Visible Then Exit Sub", initialize)
        self.assertIn("NMDC_EnsureReferenceLiveFilterControls ActiveSheet", initialize)
        activate = self.live.split("Public Sub NMDC_LiveFilterSheetActivate", 1)[1].split("End Sub", 1)[0]
        self.assertIn("NMDC_EnsureReferenceLiveFilterControls Sh", activate)

    def test_real_package_assembler_copies_feature_and_excludes_dormant_form(self):
        with tempfile.TemporaryDirectory(dir=ROOT / "tests") as temp:
            fixture = Path(temp)
            for relative in ("excel/base_chunks", "excel/vba", "config"):
                shutil.copytree(ROOT / relative, fixture / relative)
            (fixture / "dist").mkdir()
            (fixture / "dist/nmdc_index_engine.exe").write_bytes(b"test fixture")
            (fixture / "packaging").mkdir()
            for relative in (
                "packaging/Create_NMDC_Document_Index.vbs",
                "packaging/README.md",
                "tests/excel_e2e/assemble_package.py",
            ):
                destination = fixture / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(ROOT / relative, destination)

            result = subprocess.run(
                [sys.executable, str(fixture / "tests/excel_e2e/assemble_package.py")],
                cwd=fixture,
                check=False,
                capture_output=True,
                text=True,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            deployed_vba = fixture / "production-package/vba"
            deployed = (deployed_vba / "modNMDC_LiveFilter.bas").read_text(encoding="utf-8-sig")
            self.assertIn('chooseButton.Caption = "CHANGE COLUMN"', deployed)
            self.assertIn("NMDC_LiveFilterChooseColumn", deployed)
            self.assertFalse((deployed_vba / "frmNMDC_LiveFilter.frm").exists())
            self.assertFalse((deployed_vba / "frmNMDC_LiveFilter.frx").exists())

            workflow = (ROOT / ".github/workflows/production-package.yml").read_text(encoding="utf-8")
            self.assertIn("Copy-Item .\\excel\\vba\\*.bas", workflow)
            self.assertIn('Remove-Item "$package\\vba\\frmNMDC_LiveFilter.frm"', workflow)
            self.assertIn('Remove-Item "$package\\vba\\frmNMDC_LiveFilter.frx"', workflow)


if __name__ == "__main__":
    unittest.main(verbosity=2)
