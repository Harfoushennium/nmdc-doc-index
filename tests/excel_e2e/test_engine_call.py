import win32com.client
from pathlib import Path
import sys

pkg = Path('production-package').resolve()
wb_path = pkg / 'NMDC_Document_Index.xlsm'

print("Opening Excel...")
excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = True
excel.DisplayAlerts = False

try:
    print(f"Opening {wb_path}...")
    wb = excel.Workbooks.Open(str(wb_path))
    print("Testing NMDC_RunEngine with export-excel...")
    res = excel.Run('NMDC_RunEngine', 'export-excel')
    print("NMDC_RunEngine result:", res)
    wb.Close(False)
    print("Closed workbook.")
finally:
    excel.Quit()
    print("Quit Excel.")
