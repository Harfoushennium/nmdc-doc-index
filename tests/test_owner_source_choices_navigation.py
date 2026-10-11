import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class OwnerSourceChoicesNavigationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.owner_ux = (ROOT / "excel/vba/modNMDC_OwnerUX.bas").read_text(encoding="utf-8-sig")
        cls.checkboxes = (ROOT / "excel/vba/modNMDC_Checkboxes.bas").read_text(encoding="utf-8-sig")

    def test_pending_update_navigation_replaces_owned_notice_and_uses_fixed_row_two_geometry(self):
        guide = self.owner_ux.split("Private Sub NMDC_ApplyPendingUpdateDecisionGuide()", 1)[1].split("End Sub", 1)[0]
        legacy_guard = guide.split('If CStr(ws.Range("A2").Value2) = ', 1)[1].split("End If", 1)[0]
        self.assertIn('"No staged update is included in this review-only baseline. Use the macro-enabled package to create a fresh staged comparison."', legacy_guard)
        self.assertIn('ws.Range("A2:I2").UnMerge', legacy_guard)
        self.assertIn('ws.Range("A2:I2").ClearContents', legacy_guard)
        self.assertEqual(guide.count('ws.Range("A2:I2").ClearContents'), 1)
        self.assertIn('Set anchor = ws.Range("A2")', guide)
        self.assertIn("Const SOURCE_BUTTON_WIDTH As Double = 150", guide)
        self.assertIn("Const GUIDE_GAP As Double = 8", guide)
        self.assertIn("Const GUIDE_WIDTH As Double = 600", guide)
        self.assertIn('sourceButton.Name = "NMDC_Pending_SourceNavigation"', guide)
        self.assertIn('!NMDC_GoToSourceChoices', guide)
        self.assertIn('sourceButton.TextFrame.Characters.Text = "SOURCE CHOICES"', guide)
        self.assertIn("sourceButton.AlternativeText", guide)
        self.assertIn(".Placement = xlFreeFloating", guide)
        self.assertIn('guide.Name = "NMDC_Pending_SourceGuide"', guide)
        self.assertIn("anchor.Left + SOURCE_BUTTON_WIDTH + GUIDE_GAP", guide)
        self.assertNotIn('ws.Range("A4:I4").UnMerge', guide)
        self.assertIn('ws.Shapes("NMDC_Pending_SourceNavigation").Delete', guide)
        self.assertIn('ws.Shapes("NMDC_Pending_SourceGuide").Delete', guide)

    def test_action_rebuilds_then_gotos_dynamic_source_selection_header(self):
        action = self.owner_ux.split("Public Sub NMDC_GoToSourceChoices()", 1)[1].split("End Sub", 1)[0]
        self.assertIn('Set sourceTable = ws.ListObjects("SourceSelection")', action)
        self.assertIn("If sourceTable Is Nothing Then NMDC_RebuildPendingSourcePanel", action)
        self.assertNotIn("NMDC_RunEngine", action)
        self.assertNotIn("NMDC_RefreshReviewDataFast", action)
        rebuild_at = action.index("NMDC_RebuildPendingSourcePanel")
        acquire_at = action.index('Set sourceTable = ws.ListObjects("SourceSelection")', rebuild_at + 1)
        goto_at = action.index("Application.Goto Reference:=sourceTable.HeaderRowRange, Scroll:=True")
        self.assertLess(rebuild_at, acquire_at)
        self.assertLess(acquire_at, goto_at)
        self.assertNotIn("40503", action)

        legacy_open = self.owner_ux.split("Public Sub NMDC_OpenSourceSelection()", 1)[1].split("End Sub", 1)[0]
        self.assertIn('NMDC_RunEngine("export-excel")', legacy_open)
        self.assertIn("NMDC_RebuildPendingSourcePanel", legacy_open)
        self.assertIn("NMDC_GoToSourceChoices", legacy_open)

        ensure = self.checkboxes.split("Private Function NMDC_EnsurePendingSourceTable", 1)[1].split("End Function", 1)[0]
        self.assertIn('Set pendingTable = ws.ListObjects("PendingUpdate")', ensure)
        self.assertIn("headerRow = pendingTable.Range.Row + pendingTable.Range.Rows.Count + 3", ensure)
        self.assertIn('table.Name = SOURCE_PANEL_TABLE', ensure)
        self.assertIn('Private Const SOURCE_PANEL_TABLE As String = "SourceSelection"', self.checkboxes)

    def test_navigation_preserves_batched_source_choice_workflow(self):
        action = self.owner_ux.split("Public Sub NMDC_OpenSourceSelection()", 1)[1].split("End Sub", 1)[0]
        self.assertIn("NMDC_RebuildPendingSourcePanel", action)
        save_action = self.checkboxes.split("Public Sub NMDC_SaveSourceSelections()", 1)[1].split("End Sub", 1)[0]
        self.assertIn('NMDC_RunEngine("save-source-selections", args)', save_action)
        self.assertIn("NMDC_UpdateChangedFilesFast", save_action)
        self.assertIn("Save Source Choices & Restage", self.checkboxes)
        self.assertNotIn("NMDC_RunEngine", self.checkboxes.split("Public Sub NMDC_SourceCheckboxClicked()", 1)[1].split("End Sub", 1)[0])

    def test_real_package_assembler_deploys_navigation_and_excludes_dormant_form(self):
        with tempfile.TemporaryDirectory(dir=ROOT / "tests") as temp:
            fixture = Path(temp)
            for relative in ("excel/base_chunks", "excel/vba", "config"):
                shutil.copytree(ROOT / relative, fixture / relative)
            (fixture / "dist").mkdir()
            (fixture / "dist/nmdc_index_engine.exe").write_bytes(b"test fixture")
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
            deployed = (deployed_vba / "modNMDC_OwnerUX.bas").read_text(encoding="utf-8-sig")
            self.assertIn('!NMDC_GoToSourceChoices', deployed)
            self.assertIn('Set anchor = ws.Range("A2")', deployed)
            self.assertIn('ws.Range("A2:I2").UnMerge', deployed)
            self.assertIn("Application.Goto Reference:=sourceTable.HeaderRowRange, Scroll:=True", deployed)
            self.assertFalse((deployed_vba / "frmNMDC_LiveFilter.frm").exists())
            self.assertFalse((deployed_vba / "frmNMDC_LiveFilter.frx").exists())

            workflow = (ROOT / ".github/workflows/production-package.yml").read_text(encoding="utf-8")
            self.assertIn("Copy-Item .\\excel\\vba\\*.bas", workflow)
            self.assertIn('Remove-Item "$package\\vba\\frmNMDC_LiveFilter.frm"', workflow)
            self.assertIn('Remove-Item "$package\\vba\\frmNMDC_LiveFilter.frx"', workflow)


if __name__ == "__main__":
    unittest.main(verbosity=2)
