import win32gui
import win32con
import time
import argparse


def list_dialogs():
    dialogs = []
    seen = set()

    def visit(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "#32770":
            if hwnd not in seen:
                seen.add(hwnd)
                dialogs.append((hwnd, win32gui.GetWindowText(hwnd)))
        return True

    try:
        import win32service, ctypes
        user32 = ctypes.windll.user32
        hWinSta = win32service.GetProcessWindowStation()
        desktops = []
        def desk_cb(dname, _):
            desktops.append(dname)
            return True
        DESKTOPENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_wchar_p, ctypes.c_long)
        user32.EnumDesktopsW(int(hWinSta), DESKTOPENUMPROC(desk_cb), 0)
        for dname in desktops:
            try:
                hdesk = win32service.OpenDesktop(dname, 0, False, 0x0100)
                win32gui.EnumDesktopWindows(hdesk, visit, None)
            except Exception:
                pass
    except Exception:
        pass

    try:
        win32gui.EnumWindows(visit, None)
    except Exception:
        pass
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
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    closed.append(title)
                except Exception:
                    pass
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
