import win32com.client
import sys

def test_ole():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    
    # Try adding Forms.TextBox.1
    try:
        ole = ws.OLEObjects().Add(ClassType="Forms.TextBox.1", Left=100, Top=50, Width=200, Height=24)
        print("FORMS_TEXTBOX_SUCCESS: OLEObject created successfully!")
        ole.Object.Text = "Test Search"
        print(f"TEXT_SET_SUCCESS: {ole.Object.Text}")
    except Exception as ex:
        print(f"FORMS_TEXTBOX_FAILED: {ex}")
        
    wb.Close(False)
    excel.Quit()

if __name__ == "__main__":
    test_ole()
