import win32gui
import win32con
import time
import argparse


def list_dialogs():
    dialogs = []

    def visit(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "#32770":
            dialogs.append((hwnd, win32gui.GetWindowText(hwnd)))

    try:
        win32gui.EnumWindows(visit, None)
    except Exception as exc:
        print("DIALOG_ENUMERATION_BLOCKED", exc, flush=True)
    return dialogs


def child_texts(hwnd):
    values = []
    win32gui.EnumChildWindows(
        hwnd,
        lambda child, _: values.append((win32gui.GetClassName(child), win32gui.GetWindowText(child))),
        None,
    )
    return values

def close_msgbox(timeout=10):
    """Dismiss only known setup/VBA dialogs and report anything unexpected."""
    known = {"NMDC Document Index Setup", "Microsoft Excel", "Microsoft Visual Basic for Applications"}
    deadline = time.time() + timeout
    closed = []
    unexpected = []
    while time.time() < deadline:
        dialogs = list_dialogs()
        for hwnd, title in dialogs:
            print("Dialog children:", title, child_texts(hwnd), flush=True)
            if title in known or title.startswith("NMDC Document Index"):
                win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                closed.append(title)
            elif title and title not in unexpected:
                unexpected.append(title)
        if not dialogs:
            break
        time.sleep(0.2)
    print("CLOSED_DIALOGS", closed)
    print("UNEXPECTED_DIALOGS", unexpected)
    return unexpected

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--timeout", type=float, default=10)
    args = parser.parse_args()
    raise SystemExit(2 if close_msgbox(args.timeout) else 0)
