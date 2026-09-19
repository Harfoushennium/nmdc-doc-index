import win32gui
import win32con
import time

def close_msgbox():
    hwnd = win32gui.FindWindow("#32770", "NMDC Document Index Setup")
    if hwnd:
        print("Found setup MsgBox window:", hwnd)
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        print("Sent WM_CLOSE to setup MsgBox")
    else:
        print("No setup MsgBox found")

if __name__ == "__main__":
    close_msgbox()
