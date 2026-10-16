import win32com.client
import sys

def check():
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        try:
            vp = wb.VBProject
            name = vp.Name
            print(f"VBA_TRUST_OK: {name}")
        except Exception as ex:
            print(f"VBA_TRUST_FAILED: {ex}")
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"EXCEL_DISPATCH_ERROR: {e}")

if __name__ == "__main__":
    check()
