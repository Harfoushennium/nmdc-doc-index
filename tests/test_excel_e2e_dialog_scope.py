import importlib.util
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
HARNESS_PATH = ROOT / "tests" / "excel_e2e" / "test_real_excel_simulation.py"
SPEC = importlib.util.spec_from_file_location("real_excel_simulation_dialog_test", HARNESS_PATH)
HARNESS = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = HARNESS
SPEC.loader.exec_module(HARNESS)


class FakeProcess:
    args = ["cscript.exe", "setup.vbs"]
    returncode = 0

    def __init__(self):
        self.poll_count = 0

    def poll(self):
        self.poll_count += 1
        return None if self.poll_count == 1 else 0

    def communicate(self, timeout=None):
        return "final workbook saved", ""


class RunSetupDialogScopeTests(unittest.TestCase):
    def run_setup_with_dialogs(self, dialogs, owner_paths):
        # `dialogs` maps integer HWNDs to titles; `owner_paths` maps HWND/PID
        # to executable paths. An absent owner path simulates metadata failure.
        posted = []
        closed = []
        gui = types.ModuleType("win32gui")
        gui.IsWindowVisible = lambda hwnd: True
        gui.GetClassName = lambda hwnd: "#32770"
        gui.GetWindowText = lambda hwnd: dialogs[hwnd]
        enum = lambda callback, arg: [callback(hwnd, arg) for hwnd in dialogs]
        gui.EnumDesktopWindows = enum
        gui.EnumWindows = enum
        gui.PostMessage = lambda *args: posted.append(args)

        process = types.ModuleType("win32process")
        process.GetWindowThreadProcessId = lambda hwnd: (1, hwnd)
        process.GetModuleFileNameEx = lambda handle, _index: owner_paths[handle[2]]
        api = types.ModuleType("win32api")
        api.OpenProcess = lambda access, inherit, pid: (access, inherit, pid)
        api.CloseHandle = lambda handle: closed.append(handle)
        service = types.ModuleType("win32service")
        service.OpenDesktop = lambda *_args: object()
        con = types.ModuleType("win32con")
        con.WM_COMMAND, con.WM_CLOSE = 0x111, 0x10

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temp_dir:
            package = Path(temp_dir)
            (package / "Create_NMDC_Document_Index.vbs").write_text("", encoding="utf-8")
            fake_modules = {
                "win32gui": gui,
                "win32process": process,
                "win32api": api,
                "win32service": service,
                "win32con": con,
            }
            with patch.dict(sys.modules, fake_modules), \
                 patch.object(HARNESS.subprocess, "Popen", return_value=FakeProcess()), \
                 patch.object(HARNESS.time, "time", return_value=0), \
                 patch.object(HARNESS.time, "sleep", return_value=None):
                result = HARNESS._run_setup(package, timeout=5)
        return result, posted, closed

    def test_unrelated_apple_updater_does_not_fail_successful_setup(self):
        self.assertEqual(
            HARNESS._dialog_process_basename(r"C:\Apple\SoftwareUpdate.exe"),
            "softwareupdate.exe",
        )
        result, posted, _closed = self.run_setup_with_dialogs(
            {101: "Apple Software Update"}, {101: r"C:\Apple\SoftwareUpdate.exe"}
        )
        self.assertEqual(result, ("PASS", "final workbook saved"))
        self.assertEqual(posted, [])

    def test_unrelated_owner_spoofing_setup_or_excel_titles_is_never_messaged(self):
        for title in ("NMDC Document Index Setup", "Microsoft Excel"):
            with self.subTest(title=title):
                result, posted, _closed = self.run_setup_with_dialogs(
                    {101: title}, {101: r"C:\Apple\SoftwareUpdate.exe"}
                )
                self.assertEqual(result, ("PASS", "final workbook saved"))
                self.assertEqual(posted, [])

    def test_expected_excel_reset_and_cscript_setup_dialogs_keep_messages(self):
        result, posted, _closed = self.run_setup_with_dialogs(
            {101: "Reset All Records", 102: "NMDC Document Index Setup"},
            {101: r"C:\Excel\EXCEL.EXE", 102: r"C:\Windows\System32\cscript.exe"},
        )
        self.assertEqual(result, ("PASS", "final workbook saved"))
        self.assertIn((101, 0x111, 1, 0), posted)
        self.assertIn((102, 0x111, 6, 0), posted)
        self.assertIn((102, 0x111, 1, 0), posted)
        self.assertIn((102, 0x10, 0, 0), posted)
        self.assertEqual(len(posted), 4)

    def test_unexpected_excel_runtime_error_remains_failure(self):
        result, posted, _closed = self.run_setup_with_dialogs(
            {101: "Microsoft Excel - Runtime Error"}, {101: r"C:\Excel\EXCEL.EXE"}
        )
        self.assertEqual(result[0], "FAIL")
        self.assertIn("Microsoft Excel - Runtime Error", result[1])
        self.assertEqual(posted, [])

    def test_unknown_owner_runtime_error_remains_failure_and_handle_closes(self):
        result, posted, closed = self.run_setup_with_dialogs(
            {101: "Runtime error"}, {}
        )
        self.assertEqual(result[0], "FAIL")
        self.assertIn("Runtime error", result[1])
        self.assertEqual(posted, [])
        self.assertEqual(len(closed), 1)

    def test_unknown_owner_expected_setup_prompt_is_retained_and_handled(self):
        result, posted, closed = self.run_setup_with_dialogs(
            {101: "NMDC Document Index Setup"}, {}
        )
        self.assertEqual(result, ("PASS", "final workbook saved"))
        self.assertTrue(posted)
        self.assertEqual(len(closed), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
