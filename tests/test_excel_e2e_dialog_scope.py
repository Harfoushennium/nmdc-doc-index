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


class UndoPromptAuthorizationTests(unittest.TestCase):
    def confirmable(self, *, macro=HARNESS.UNDO_MACRO, verified=True, pid=4321,
                    dialog_pid=4321, title=HARNESS.UNDO_TITLE,
                    text=HARNESS.UNDO_PROMPT_TEXT):
        return HARNESS._undo_prompt_is_confirmable(
            macro, verified, pid, dialog_pid, title, text
        )

    def test_intended_excel_and_verified_negative_state_allows_yes_only(self):
        posted = []
        allowed = HARNESS._confirm_undo_prompt(
            77, HARNESS.UNDO_MACRO, True, 4321, 4321,
            HARNESS.UNDO_TITLE, HARNESS.UNDO_PROMPT_TEXT,
            lambda *args: posted.append(args), 0x111,
        )
        self.assertTrue(allowed)
        self.assertEqual(posted, [(77, 0x111, 6, 0)])

    def test_realistic_message_box_extracts_body_not_button_captions(self):
        children = {1: ("Static", HARNESS.UNDO_PROMPT_TEXT.replace("version? ", "version?\r\n\r\n")),
                    2: ("Button", "&Yes"), 3: ("Button", "&No"),
                    4: ("Static", "")}
        gui = types.SimpleNamespace(
            EnumChildWindows=lambda _hwnd, callback, arg: [callback(child, arg) for child in children],
            GetClassName=lambda child: children[child][0],
            GetWindowText=lambda child: children[child][1],
        )
        old_all_child_text = " ".join(value for _class_name, value in children.values())
        self.assertFalse(HARNESS._undo_prompt_is_confirmable(
            HARNESS.UNDO_MACRO, True, 4321, 4321,
            HARNESS.UNDO_TITLE, old_all_child_text,
        ))
        extracted = HARNESS._dialog_message_text(77, gui)
        posted = []
        allowed = HARNESS._confirm_undo_prompt(
            77, HARNESS.UNDO_MACRO, True, 4321, 4321,
            HARNESS.UNDO_TITLE, extracted, lambda *args: posted.append(args), 0x111,
        )
        self.assertTrue(allowed)
        self.assertEqual(posted, [(77, 0x111, 6, 0)])

    def test_extra_static_prompt_text_never_receives_yes(self):
        children = {1: ("Static", HARNESS.UNDO_PROMPT_TEXT),
                    2: ("Static", "An unrelated warning")}
        gui = types.SimpleNamespace(
            EnumChildWindows=lambda _hwnd, callback, arg: [callback(child, arg) for child in children],
            GetClassName=lambda child: children[child][0],
            GetWindowText=lambda child: children[child][1],
        )
        self.assert_no_undo_message(text=HARNESS._dialog_message_text(77, gui))

    def test_expected_refusal_gets_idok_only_from_matching_excel(self):
        children = {1: ("Static", HARNESS.UNDO_REFUSAL_TEXT.replace("restore, or", "restore,\r\nor")),
                    2: ("Button", "OK"), 3: ("Static", "")}
        gui = types.SimpleNamespace(
            EnumChildWindows=lambda _hwnd, callback, arg: [callback(child, arg) for child in children],
            GetClassName=lambda child: children[child][0],
            GetWindowText=lambda child: children[child][1],
        )
        posted = []
        handled = HARNESS._acknowledge_undo_refusal(
            78, HARNESS.UNDO_MACRO, True, 4321, 4321,
            HARNESS.UNDO_REFUSAL_TITLE, HARNESS._dialog_message_text(78, gui),
            lambda *args: posted.append(args), 0x111,
        )
        self.assertTrue(handled)
        self.assertEqual(posted, [(78, 0x111, 1, 0)])

    def test_refusal_unknown_owner_other_macro_or_changed_prompt_is_untouched(self):
        cases = (
            (HARNESS.UNDO_MACRO, True, 4321, None, HARNESS.UNDO_REFUSAL_TITLE, HARNESS.UNDO_REFUSAL_TEXT),
            (HARNESS.UNDO_MACRO, True, 4321, 9876, HARNESS.UNDO_REFUSAL_TITLE, HARNESS.UNDO_REFUSAL_TEXT),
            ("NMDC_FullRescan", True, 4321, 4321, HARNESS.UNDO_REFUSAL_TITLE, HARNESS.UNDO_REFUSAL_TEXT),
            (HARNESS.UNDO_MACRO, False, 4321, 4321, HARNESS.UNDO_REFUSAL_TITLE, HARNESS.UNDO_REFUSAL_TEXT),
            (HARNESS.UNDO_MACRO, True, 4321, 4321, "Microsoft Excel", HARNESS.UNDO_REFUSAL_TEXT),
            (HARNESS.UNDO_MACRO, True, 4321, 4321, HARNESS.UNDO_REFUSAL_TITLE, "Different refusal"),
        )
        for macro, verified, intended, owner, title, text in cases:
            with self.subTest(macro=macro, owner=owner, title=title):
                posted = []
                handled = HARNESS._acknowledge_undo_refusal(
                    78, macro, verified, intended, owner, title, text,
                    lambda *args: posted.append(args), 0x111,
                )
                self.assertFalse(handled)
                self.assertEqual(posted, [])

    def test_prior_approval_state_refuses_confirmation(self):
        class Cell:
            def __init__(self, value):
                self.Value = value

        class Cells:
            def __init__(self, values):
                self.values = values

            def __call__(self, row, column):
                return Cell(self.values[column - 1])

        class Row:
            def __init__(self, values):
                self.Range = types.SimpleNamespace(Cells=Cells(values))

        class Rows:
            Count = 1

            def __call__(self, _index):
                return Row(("Runtime Folder", "runtime"))

        class Table:
            ListRows = Rows()

        class Workbook:
            def __init__(self, fullname):
                self.FullName = str(fullname)

            def Worksheets(self, _name):
                return types.SimpleNamespace(ListObjects=lambda _table: Table())

        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as temp_dir:
            package = Path(temp_dir) / "scratch_package"
            runtime = package / "runtime" / "approved"
            workbook = Workbook(package / "NMDC_Document_Index.xlsm")
            allowed, _detail = HARNESS._undo_negative_state_guard(
                workbook, allowed_roots=(Path(temp_dir),)
            )
            self.assertTrue(allowed)
            self.assertTrue(self.confirmable(verified=allowed))
            (runtime / "versions").mkdir(parents=True)
            (runtime / "versions" / "APP-OLD.json").write_text("{}", encoding="utf-8")
            allowed, _detail = HARNESS._undo_negative_state_guard(
                workbook, allowed_roots=(Path(temp_dir),)
            )
            self.assertFalse(allowed)
            self.assert_no_undo_message(verified=allowed)

    def test_current_pointer_refuses_confirmation(self):
        workbook, allowed_roots, runtime, _temp_dir = self.make_guard_fixture()
        approved = runtime / "approved"
        approved.mkdir(parents=True)
        (approved / "current.json").write_text("{}", encoding="utf-8")
        allowed, _detail = HARNESS._undo_negative_state_guard(
            workbook, allowed_roots=allowed_roots
        )
        self.assertFalse(allowed)
        self.assert_no_undo_message(verified=allowed)

    def make_guard_fixture(self):
        # Keep the TemporaryDirectory alive through the caller's guard check.
        temp_dir = tempfile.TemporaryDirectory(dir=Path(__file__).parent)
        self.addCleanup(temp_dir.cleanup)
        package = Path(temp_dir.name) / "scratch_package"
        runtime = package / "runtime"

        class Cell:
            def __init__(self, value):
                self.Value = value

        class Cells:
            def __call__(self, _row, column):
                return Cell(("Runtime Folder", "runtime")[column - 1])

        class Row:
            Range = types.SimpleNamespace(Cells=Cells())

        class Rows:
            Count = 1

            def __call__(self, _index):
                return Row()

        class Table:
            ListRows = Rows()

        class Workbook:
            FullName = str(package / "NMDC_Document_Index.xlsm")

            def Worksheets(self, _name):
                return types.SimpleNamespace(ListObjects=lambda _table: Table())

        return Workbook(), (Path(temp_dir.name),), runtime, temp_dir

    def test_runtime_outside_package_refuses_confirmation(self):
        workbook, allowed_roots, runtime, _temp_dir = self.make_guard_fixture()
        class BadRuntimeWorkbook:
            FullName = workbook.FullName

            def Worksheets(self, _name):
                class Cell:
                    def __init__(self, value):
                        self.Value = value

                class Cells:
                    def __call__(self, _row, column):
                        return Cell(("Runtime Folder", str(Path("..") / "outside-runtime"))[column - 1])

                class Row:
                    Range = types.SimpleNamespace(Cells=Cells())

                class Rows:
                    Count = 1

                    def __call__(self, _index):
                        return Row()

                return types.SimpleNamespace(ListObjects=lambda _table: types.SimpleNamespace(ListRows=Rows()))

        allowed, detail = HARNESS._undo_negative_state_guard(
            BadRuntimeWorkbook(), allowed_roots=allowed_roots
        )
        self.assertFalse(allowed)
        self.assertIn("not package-local", detail)
        self.assertFalse(self.confirmable(verified=allowed))

    def test_no_prior_undo_invocation_is_reported_blocked_not_pass(self):
        simulation = HARNESS.ExcelSimulation.__new__(HARNESS.ExcelSimulation)
        simulation.report_dir = Path(__file__).parent
        simulation.excel = object()
        simulation.results = []
        invoked = []
        simulation._macro = lambda name: invoked.append(name) or (
            "COM invoked Undo; expected no-prior refusal observed and OK sent"
        )
        simulation._observe_dialogs = lambda: []
        with patch.object(HARNESS, "_capture_screen", return_value="mock.png"):
            simulation._record(13, "Undo", simulation._undo_test)
        self.assertEqual(invoked, [HARNESS.UNDO_MACRO])
        self.assertEqual(simulation.results[0].status, "BLOCKED")
        self.assertIn("expected no-prior refusal dialog was observed", simulation.results[0].detail)
        self.assertIn("does not independently establish the engine cause", simulation.results[0].detail)
        self.assertIn("positive Undo restoration", simulation.results[0].detail)

        no_dialog = HARNESS.ExcelSimulation.__new__(HARNESS.ExcelSimulation)
        no_dialog.report_dir = Path(__file__).parent
        no_dialog.excel = object()
        no_dialog.results = []
        no_dialog._macro = lambda _name: "COM invoked Undo"
        no_dialog._observe_dialogs = lambda: []
        with patch.object(HARNESS, "_capture_screen", return_value="mock.png"):
            no_dialog._record(13, "Undo", no_dialog._undo_test)
        self.assertEqual(no_dialog.results[0].status, "BLOCKED")
        self.assertIn("expected no-prior refusal dialog was not observed", no_dialog.results[0].detail)

    def test_non_disposable_workbook_is_rejected(self):
        workbook = types.SimpleNamespace(FullName=str(Path(__file__).resolve()))
        allowed, detail = HARNESS._undo_negative_state_guard(workbook)
        self.assertFalse(allowed)
        self.assertIn("outside disposable", detail)

    def test_unknown_owner_never_receives_yes(self):
        self.assert_no_undo_message(dialog_pid=None)

    def test_other_process_never_receives_yes(self):
        self.assert_no_undo_message(dialog_pid=9876)

    def test_other_macro_never_receives_yes(self):
        self.assert_no_undo_message(macro="NMDC_FullRescan")

    def test_nonmatching_prompt_never_receives_yes(self):
        self.assert_no_undo_message(title="Microsoft Excel")
        self.assert_no_undo_message(text="Unexpected confirmation")

    def assert_no_undo_message(self, **overrides):
        values = dict(
            macro_name=HARNESS.UNDO_MACRO, negative_state_verified=True,
            intended_pid=4321, dialog_pid=4321, title=HARNESS.UNDO_TITLE,
            text=HARNESS.UNDO_PROMPT_TEXT,
        )
        keys = {
            "macro": "macro_name", "verified": "negative_state_verified",
            "pid": "intended_pid", "dialog_pid": "dialog_pid",
            "title": "title", "text": "text",
        }
        values.update({keys[key]: value for key, value in overrides.items()})
        posted = []
        allowed = HARNESS._confirm_undo_prompt(
            77, values["macro_name"], values["negative_state_verified"],
            values["intended_pid"], values["dialog_pid"], values["title"],
            values["text"], lambda *args: posted.append(args), 0x111,
        )
        self.assertFalse(allowed)
        self.assertEqual(posted, [])

if __name__ == "__main__":
    unittest.main(verbosity=2)
