Attribute VB_Name = "modNMDC_Refresh"
Option Explicit

Public Sub NMDC_RefreshExchangeData()
    On Error GoTo Handler

    NMDC_LoadCsvToSheet NMDC_ExchangePath() & "\master_documents.csv", "Master Documents"
    NMDC_LoadCsvToSheet NMDC_ExchangePath() & "\revisions.csv", "Revisions"
    NMDC_LoadCsvToSheet NMDC_ExchangePath() & "\events.csv", "Transactions"
    NMDC_LoadCsvToSheet NMDC_ExchangePath() & "\pending_update.csv", "Pending Update"
    NMDC_LoadCsvToSheet NMDC_ExchangePath() & "\flags.csv", "Review Flags"
    NMDC_LoadCsvToSheet NMDC_ExchangePath() & "\history.csv", "Update History"
    NMDC_LoadErrorCsvPreserveLocal NMDC_ExchangePath() & "\errors.csv"
    NMDC_LoadDashboard NMDC_ExchangePath() & "\dashboard.csv"

    Exit Sub
Handler:
    NMDC_LogError "EXCEL_REFRESH_ERROR", _
        "Excel could not refresh one or more NMDC Index tables.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LoadCsvToSheet(ByVal csvPath As String, ByVal sheetName As String)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim qt As QueryTable

    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Len(Dir$(csvPath)) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    ws.Cells.ClearContents

    For Each qt In ws.QueryTables
        qt.Delete
    Next qt

    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFilePlatform = 65001
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    NMDC_FormatDataSheet ws
    Application.ScreenUpdating = True
    Exit Sub

Handler:
    Application.ScreenUpdating = True
    NMDC_LogError "CSV_IMPORT_ERROR", _
        "Excel could not load " & sheetName & ".", _
        "File: " & csvPath & " | " & Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LoadErrorCsvPreserveLocal(ByVal csvPath As String)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim localRows As Collection
    Dim rowValues As Variant
    Dim r As Long
    Dim c As Long
    Dim lastRow As Long
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("Error Log")

    ' If the engine did not produce an error exchange file, leave the existing log untouched.
    If Len(Dir$(csvPath)) = 0 Then Exit Sub

    Set localRows = New Collection

    ' Preserve only locally-created Excel/VBA entries. Engine-exported rows can be safely refreshed.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        For r = 2 To lastRow
            If Left$(CStr(ws.Cells(r, 3).Value), 6) = "EXCEL:" Then
                ReDim rowValues(1 To 8)
                For c = 1 To 8
                    rowValues(c) = ws.Cells(r, c).Value
                Next c
                localRows.Add rowValues
            End If
        Next r
    End If

    Application.ScreenUpdating = False
    ws.Cells.ClearContents
    For Each qt In ws.QueryTables
        qt.Delete
    Next qt

    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFilePlatform = 65001
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    For Each rowValues In localRows
        nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2
        For c = 1 To 8
            ws.Cells(nextRow, c).Value = rowValues(c)
        Next c
    Next rowValues

    NMDC_FormatDataSheet ws
    Application.ScreenUpdating = True
    Exit Sub

Handler:
    Application.ScreenUpdating = True
    NMDC_LogError "ERROR_LOG_REFRESH_ERROR", _
        "Excel could not refresh the Error Log safely.", _
        "File: " & csvPath & " | " & Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_FormatDataSheet(ByVal ws As Worksheet)
    On Error Resume Next
    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Rows(1).AutoFilter
    ws.Cells.EntireColumn.AutoFit
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Public Sub NMDC_LoadDashboard(ByVal csvPath As String)
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim temp As Worksheet
    Dim qt As QueryTable

    If Len(Dir$(csvPath)) = 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Home")
    Set temp = ThisWorkbook.Worksheets("System Data")

    temp.Cells.ClearContents
    For Each qt In temp.QueryTables
        qt.Delete
    Next qt

    Set qt = temp.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=temp.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFilePlatform = 65001
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    ' Home cells intentionally use simple fixed positions so the workbook remains easy to maintain.
    ws.Range("B5").Value = NMDC_SystemValue(temp, "Approved Status")
    ws.Range("B6").Value = NMDC_SystemValue(temp, "Approved Run ID")
    ws.Range("B7").Value = NMDC_SystemValue(temp, "Current Data Folder")
    ws.Range("E5").Value = NMDC_SystemValue(temp, "Approved Documents")
    ws.Range("E6").Value = NMDC_SystemValue(temp, "Approved Revisions")
    ws.Range("E7").Value = NMDC_SystemValue(temp, "Approved Transactions")
    ws.Range("H5").Value = NMDC_SystemValue(temp, "Pending Status")
    ws.Range("H6").Value = NMDC_SystemValue(temp, "Pending Run ID")
    ws.Range("H7").Value = "Added " & NMDC_SystemValue(temp, "Pending Added") & _
                           " | Modified " & NMDC_SystemValue(temp, "Pending Modified") & _
                           " | Removed " & NMDC_SystemValue(temp, "Pending Removed")
    ws.Range("K5").Value = NMDC_SystemValue(temp, "Review Flags")
    ws.Range("K6").Value = NMDC_SystemValue(temp, "Conflict Flags")
    Exit Sub

Handler:
    NMDC_LogError "DASHBOARD_REFRESH_ERROR", _
        "Excel could not refresh the Home dashboard.", _
        Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_SystemValue(ByVal ws As Worksheet, ByVal headerName As String) As String
    Dim hit As Range
    Set hit = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If hit Is Nothing Then
        NMDC_SystemValue = ""
    Else
        NMDC_SystemValue = CStr(ws.Cells(2, hit.Column).Value)
    End If
End Function
