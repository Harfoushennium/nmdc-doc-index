import win32com.client
from pathlib import Path

pkg = Path('production-package').resolve()
wb_path = pkg / 'NMDC_Document_Index.xlsm'

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = True
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(str(wb_path))
    print("Testing kernel32 Sleep...")
    comp = wb.VBProject.VBComponents.Add(1)
    vba_lines = [
        '#If VBA7 Then',
        'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)',
        '#Else',
        'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)',
        '#End If',
        'Public Sub TestSleep()',
        '    Sleep 100',
        'End Sub'
    ]
    comp.CodeModule.AddFromString('\n'.join(vba_lines))
    excel.Run('TestSleep')
    print("Sleep passed successfully!")
    wb.Close(False)
finally:
    excel.Quit()
