Attribute VB_Name = "modNMDC_Refresh"
Option Explicit

Public Function NMDC_RefreshExchangeData() As Boolean
    On Error GoTo Handler

    Dim refreshOk As Boolean
    refreshOk = True

    If Not NMDC_LoadCsvToSheet(NMDC_ExchangePath() & "\master_documents.csv", "Master Documents") Then refreshOk = False
    If Not NMDC_LoadCsvToSheet(NMDC_ExchangePath() & "\revisions.csv", "Revisions") Then refreshOk = False
    If Not NMDC_LoadCsvToSheet(NMDC_ExchangePath() & "\events.csv", "Transactions") Then refreshOk = False
    If Not NMDC_LoadCsvToSheet(NMDC_ExchangePath() & "\pending_update.csv", "Pending Update") Then refreshOk = False
    If Not NMDC_LoadCsvToSheet(NMDC_ExchangePath() & "\flags.csv", "Review Flags") Then refreshOk = False
    If Not NMDC_LoadCsvToSheet(NMDC_ExchangePath() & "\history.csv", "Update History") Then refreshOk = False
    If Not NMDC_LoadErrorCsvPreserveLocal(NMDC_ExchangePath() & "\errors.csv") Then refreshOk = False
    If Not NMDC_LoadDashboard(NMDC_ExchangePath() & "\dashboard.csv") Then refreshOk = False

    NMDC_RefreshExchangeData = refreshOk
    Exit Function
Handler:
    NMDC_LogError "EXCEL_REFRESH_ERROR", _
        "Excel could not refresh one or more NMDC Index tables.", _
        Err.Number & " - " & Err.Description
    NMDC_RefreshExchangeData = False
End Function

Public Function NMDC_LoadCsvToSheet(ByVal csvPath As String, ByVal sheetName As String) As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim qt As QueryTable

    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Len(Dir$(csvPath)) = 0 Then
        NMDC_LogError "CSV_MISSING", "Excel could not find the exported data for " & sheetName & ".", "File: " & csvPath
        NMDC_LoadCsvToSheet = False
        Exit Function
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
        .TextFileColumnDataTypes = NMDC_TextColumnTypes(csvPath)
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    NMDC_FormatDataSheet ws
    Application.ScreenUpdating = True
    NMDC_LoadCsvToSheet = True
    Exit Function

Handler:
    Application.ScreenUpdating = True
    NMDC_LogError "CSV_IMPORT_ERROR", _
        "Excel could not load " & sheetName & ".", _
        "File: " & csvPath & " | " & Err.Number & " - " & Err.Description
    NMDC_LoadCsvToSheet = False
End Function

Public Function NMDC_LoadErrorCsvPreserveLocal(ByVal csvPath As String) As Boolean
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
    If Len(Dir$(csvPath)) = 0 Then
        NMDC_LogError "CSV_MISSING", "Excel could not find the exported Error Log.", "File: " & csvPath
        NMDC_LoadErrorCsvPreserveLocal = False
        Exit Function
    End If

    Set localRows = New Collection

    ' Preserve only locally-created Excel/VBA entries. Engine-exported rows can be safely refreshed.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        For r = 2 To lastRow
            If Left$(CStr(ws.Cells(r, 3).Value), 6) = "EXCEL:" Then
                ReDim rowValues(1 To 10)
                For c = 1 To 10
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
        .TextFileColumnDataTypes = NMDC_TextColumnTypes(csvPath)
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    For Each rowValues In localRows
        nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2
        For c = 1 To 10
            ws.Cells(nextRow, c).Value = rowValues(c)
        Next c
    Next rowValues

    NMDC_FormatDataSheet ws
    Application.ScreenUpdating = True
    NMDC_LoadErrorCsvPreserveLocal = True
    Exit Function

Handler:
    Application.ScreenUpdating = True
    NMDC_LogError "ERROR_LOG_REFRESH_ERROR", _
        "Excel could not refresh the Error Log safely.", _
        "File: " & csvPath & " | " & Err.Number & " - " & Err.Description
    NMDC_LoadErrorCsvPreserveLocal = False
End Function

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

Public Function NMDC_LoadDashboard(ByVal csvPath As String) As Boolean
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim temp As Worksheet
    Dim qt As QueryTable

    If Len(Dir$(csvPath)) = 0 Then
        NMDC_LogError "CSV_MISSING", "Excel could not find the exported dashboard data.", "File: " & csvPath
        NMDC_LoadDashboard = False
        Exit Function
    End If
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
        .TextFileColumnDataTypes = NMDC_TextColumnTypes(csvPath)
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    ' Home cells intentionally use simple fixed positions so the workbook remains easy to maintain.
    ws.Range("B5").Value = NMDC_SystemValue(temp, "Approved Status")
    ws.Range("B6").Value = NMDC_SystemValue(temp, "Approved Run ID")
    ws.Range("B7").Value = NMDC_SystemValue(temp, "Current Data Folder")
    ws.Range("B8").Value = NMDC_SystemValue(temp, "Last Successful Update")
    ws.Range("E5").Value = NMDC_SystemValue(temp, "Approved Documents")
    ws.Range("E6").Value = NMDC_SystemValue(temp, "Approved Revisions")
    ws.Range("E7").Value = NMDC_SystemValue(temp, "Approved Transactions")
    ws.Range("H5").Value = NMDC_SystemValue(temp, "Pending Status")
    ws.Range("H6").Value = NMDC_SystemValue(temp, "Pending Run ID")
    ws.Range("H7").Value = "Added " & NMDC_SystemValue(temp, "Pending Added") & _
                           " | Modified " & NMDC_SystemValue(temp, "Pending Modified") & _
                           " | Removed " & NMDC_SystemValue(temp, "Pending Removed") & _
                           " | Unchanged " & NMDC_SystemValue(temp, "Pending Unchanged")
    ws.Range("K5").Value = NMDC_SystemValue(temp, "Review Flags")
    ws.Range("K6").Value = NMDC_SystemValue(temp, "Conflict Flags")
    NMDC_LoadDashboard = True
    Exit Function

Handler:
    NMDC_LogError "DASHBOARD_REFRESH_ERROR", _
        "Excel could not refresh the Home dashboard.", _
        Err.Number & " - " & Err.Description
    NMDC_LoadDashboard = False
End Function

Private Function NMDC_TextColumnTypes(ByVal csvPath As String) As Variant
    On Error GoTo Handler

    Dim fileNumber As Integer
    Dim headerLine As String
    Dim headerFields As Variant
    Dim dataTypes() As Integer
    Dim index As Long

    fileNumber = FreeFile
    Open csvPath For Input As #fileNumber
    Line Input #fileNumber, headerLine
    Close #fileNumber

    ' Exchange headers are controlled by the engine and contain no commas.
    ' Import every field as text so Excel cannot change revisions such as 00,
    ' document identifiers with leading zeroes, or references such as 1-2.
    headerFields = Split(headerLine, ",")
    ReDim dataTypes(0 To UBound(headerFields))
    For index = 0 To UBound(dataTypes)
        dataTypes(index) = xlTextFormat
    Next index
    NMDC_TextColumnTypes = dataTypes
    Exit Function

Handler:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    NMDC_TextColumnTypes = Array(xlTextFormat)
End Function

Private Function NMDC_SystemValue(ByVal ws As Worksheet, ByVal headerName As String) As String
    Dim hit As Range
    Set hit = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If hit Is Nothing Then
        NMDC_SystemValue = ""
    Else
        NMDC_SystemValue = CStr(ws.Cells(2, hit.Column).Value)
    End If
End Function
