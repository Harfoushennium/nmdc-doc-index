import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class FullRefreshSourcePanelTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.refresh = (ROOT / "excel/vba/modNMDC_Refresh.bas").read_text(encoding="utf-8-sig")
        cls.procedure = cls.refresh.split("Public Function NMDC_RefreshExchangeData() As Boolean", 1)[1].split(
            "End Function", 1
        )[0]

    def test_full_refresh_imports_then_builds_and_verifies_source_panel(self):
        procedure = self.procedure
        imports = (
            'NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\\master_documents.csv"',
            'NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\\revisions.csv"',
            'NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\\events.csv"',
            'NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\\pending_update.csv"',
            'NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\\flags.csv"',
            'NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\\history.csv"',
            'NMDC_LoadErrorCsvPreserveLocal(NMDC_ExchangePath() & "\\errors.csv")',
        )
        positions = [procedure.index(call) for call in imports]
        self.assertEqual(positions, sorted(positions))

        guidance = procedure.index("NMDC_ApplyWorkbookGuidance")
        owner_ui = procedure.index("NMDC_ApplyOwnerUX")
        rebuild = procedure.index("NMDC_RebuildPendingSourcePanel")
        lookup = procedure.index('Set sourceSelectionTable = pendingSheet.ListObjects("FileReview")')
        missing_guard = procedure.index("If sourceSelectionTable Is Nothing Then")
        clean_exit = procedure.index("CleanExit:")
        self.assertLess(guidance, owner_ui)
        self.assertLess(owner_ui, rebuild)
        dashboard = procedure.index('NMDC_LoadDashboard(NMDC_ExchangePath() & "\\dashboard.csv")')
        self.assertLess(rebuild, dashboard)
        self.assertLess(dashboard, lookup)
        self.assertLess(lookup, missing_guard)
        self.assertLess(missing_guard, clean_exit)

    def test_missing_source_panel_marks_full_refresh_failed(self):
        procedure = self.procedure
        missing_guard = procedure.split("If sourceSelectionTable Is Nothing Then", 1)[1].split("End If", 1)[0]
        self.assertIn("refreshOk = False", missing_guard)
        self.assertIn('NMDC_LogError "SOURCE_SELECTION_REFRESH_ERROR"', missing_guard)
        self.assertIn("NMDC_RefreshExchangeData = refreshOk", procedure)
        self.assertIn("refreshOk = True", procedure)


if __name__ == "__main__":
    unittest.main(verbosity=2)
