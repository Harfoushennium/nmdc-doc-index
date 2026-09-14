Attribute VB_Name = "modNMDC_Refresh"
Option Explicit

Public Function NMDC_RefreshExchangeData() As Boolean
    On Error GoTo Handler

    Dim refreshOk As Boolean
    refreshOk = True

    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\master_documents.csv", "Master Documents", "MasterDocuments") Then refreshOk = False
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\revisions.csv", "Revisions", "RevisionRegister") Then refreshOk = False
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\events.csv", "Transactions", "EventRegister") Then refreshOk = False
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\pending_update.csv", "Pending Update", "PendingUpdate") Then refreshOk = False
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\flags.csv", "Review Flags", "ReviewFlags") Then refreshOk = False
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\history.csv", "Update History", "UpdateHistory") Then refreshOk = False
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

Public Function NMDC_LoadCsvToTable(ByVal csvPath As String, ByVal sheetName As String, ByVal tableName As String) As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim temp As Worksheet
    Dim imported As Range
    Dim targetRange As Range
    Dim dataCount As Long
    Dim tableRows As Long
    Dim columnCount As Long
    Dim headerRow As Long
    Dim firstColumn As Long

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set table = Nothing
    On Error Resume Next
    Set table = ws.ListObjects(tableName)
    On Error GoTo Handler

    If table Is Nothing Then
        NMDC_LogError "TABLE_MISSING", _
            "Excel could not find the required table on " & sheetName & ".", _
            "Expected table: " & tableName
        NMDC_LoadCsvToTable = False
        Exit Function
    End If

    If Not NMDC_FileExists(csvPath) Then
        NMDC_LogError "CSV_MISSING", "Excel could not find the exported data for " & sheetName & ".", "File: " & csvPath
        NMDC_LoadCsvToTable = False
        Exit Function
    End If

    Application.ScreenUpdating = False
    If Not NMDC_ImportCsvToTemporarySheet(csvPath, temp, imported) Then GoTo ImportFailed

    columnCount = imported.Columns.Count
    dataCount = imported.Rows.Count - 1
    If dataCount < 0 Then dataCount = 0
    tableRows = dataCount
    If tableRows < 1 Then tableRows = 1

    headerRow = table.HeaderRowRange.Row
    firstColumn = table.Range.Column

    If Not table.DataBodyRange Is Nothing Then
        On Error Resume Next
        table.DataBodyRange.Hyperlinks.Delete
        table.DataBodyRange.ClearContents
        On Error GoTo Handler
    End If

    Set targetRange = ws.Range(ws.Cells(headerRow, firstColumn), _
                               ws.Cells(headerRow + tableRows, firstColumn + columnCount - 1))
    table.Resize targetRange
    table.HeaderRowRange.Value2 = imported.Rows(1).Value2

    If dataCount > 0 Then
        table.DataBodyRange.Value2 = imported.Offset(1, 0).Resize(dataCount, columnCount).Value2
    Else
        table.DataBodyRange.ClearContents
    End If

    NMDC_ApplyTypedFormatting table
    NMDC_ActivateDocumentLinks table

    NMDC_DeleteTemporarySheet temp
    Application.ScreenUpdating = True
    NMDC_LoadCsvToTable = True
    Exit Function

ImportFailed:
    NMDC_DeleteTemporarySheet temp
    Application.ScreenUpdating = True
    NMDC_LoadCsvToTable = False
    Exit Function

Handler:
    NMDC_DeleteTemporarySheet temp
    Application.ScreenUpdating = True
    NMDC_LogError "CSV_IMPORT_ERROR", _
        "Excel could not load " & sheetName & ".", _
        "Table: " & tableName & " | File: " & csvPath & " | " & Err.Number & " - " & Err.Description
    NMDC_LoadCsvToTable = False
End Function

Private Function NMDC_ImportCsvToTemporarySheet(ByVal csvPath As String, ByRef temp As Worksheet, ByRef imported As Range) As Boolean
    On Error GoTo Handler

    Dim qt As QueryTable
    Set temp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    temp.Visible = xlSheetVeryHidden

    Set qt = temp.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=temp.Range("A1"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFilePlatform = 65001
        .TextFileColumnDataTypes = NMDC_ColumnTypes(csvPath)
        .Refresh BackgroundQuery:=False
    End With
    Set imported = qt.ResultRange
    qt.Delete

    NMDC_ImportCsvToTemporarySheet = True
    Exit Function
Handler:
    NMDC_LogError "CSV_TEMP_IMPORT_ERROR", _
        "Excel could not read an exported CSV file.", _
        "File: " & csvPath & " | " & Err.Number & " - " & Err.Description
    NMDC_ImportCsvToTemporarySheet = False
End Function

Private Sub NMDC_DeleteTemporarySheet(ByRef temp As Worksheet)
    On Error Resume Next
    If Not temp Is Nothing Then
        Application.DisplayAlerts = False
        temp.Delete
        Application.DisplayAlerts = True
    End If
    Set temp = Nothing
End Sub

Public Function NMDC_LoadErrorCsvPreserveLocal(ByVal csvPath As String) As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim localRows As Collection
    Dim rowValues As Variant
    Dim row As ListRow
    Dim actionColumn As Long
    Dim columnIndex As Long
    Dim targetRow As ListRow

    Set ws = ThisWorkbook.Worksheets("Error Log")
    Set table = ws.ListObjects("ErrorLog")
    Set localRows = New Collection

    On Error Resume Next
    actionColumn = table.ListColumns("Action").Index
    On Error GoTo Handler
    If actionColumn = 0 Then actionColumn = 3

    For Each row In table.ListRows
        If Left$(CStr(row.Range.Cells(1, actionColumn).Value), 6) = "EXCEL:" Then
            ReDim rowValues(1 To table.ListColumns.Count)
            For columnIndex = 1 To table.ListColumns.Count
                rowValues(columnIndex) = row.Range.Cells(1, columnIndex).Value
            Next columnIndex
            localRows.Add rowValues
        End If
    Next row

    If Not NMDC_LoadCsvToTable(csvPath, "Error Log", "ErrorLog") Then
        NMDC_LoadErrorCsvPreserveLocal = False
        Exit Function
    End If

    Set table = ws.ListObjects("ErrorLog")
    For Each rowValues In localRows
        Set targetRow = NMDC_BlankOrNewRow(table)
        For columnIndex = 1 To table.ListColumns.Count
            targetRow.Range.Cells(1, columnIndex).Value = rowValues(columnIndex)
        Next columnIndex
    Next rowValues
    NMDC_ApplyTypedFormatting table

    NMDC_LoadErrorCsvPreserveLocal = True
    Exit Function
Handler:
    NMDC_LogError "ERROR_LOG_REFRESH_ERROR", _
        "Excel could not refresh the Error Log safely.", _
        "File: " & csvPath & " | " & Err.Number & " - " & Err.Description
    NMDC_LoadErrorCsvPreserveLocal = False
End Function

Private Function NMDC_BlankOrNewRow(ByVal table As ListObject) As ListRow
    If table.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(table.ListRows(1).Range) = 0 Then
            Set NMDC_BlankOrNewRow = table.ListRows(1)
            Exit Function
        End If
    End If
    Set NMDC_BlankOrNewRow = table.ListRows.Add
End Function

Private Function NMDC_ColumnTypes(ByVal csvPath As String) As Variant
    On Error GoTo Handler

    Dim fileNumber As Integer
    Dim headerLine As String
    Dim headerFields As Variant
    Dim dataTypes() As Integer
    Dim index As Long
    Dim headerName As String

    fileNumber = FreeFile
    Open csvPath For Input As #fileNumber
    Line Input #fileNumber, headerLine
    Close #fileNumber

    headerFields = Split(headerLine, ",")
    ReDim dataTypes(0 To UBound(headerFields))
    For index = 0 To UBound(headerFields)
        headerName = NMDC_CleanHeader(CStr(headerFields(index)))
        If NMDC_IsDateHeader(headerName) Or NMDC_IsNumericHeader(headerName) Then
            dataTypes(index) = xlGeneralFormat
        Else
            dataTypes(index) = xlTextFormat
        End If
    Next index
    NMDC_ColumnTypes = dataTypes
    Exit Function
Handler:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    NMDC_ColumnTypes = Array(xlTextFormat)
End Function

Private Function NMDC_CleanHeader(ByVal rawHeader As String) As String
    Dim value As String
    value = Trim$(rawHeader)
    If Left$(value, 1) = Chr$(34) Then value = Mid$(value, 2)
    If Right$(value, 1) = Chr$(34) Then value = Left$(value, Len(value) - 1)
    NMDC_CleanHeader = Replace(value, Chr$(34) & Chr$(34), Chr$(34))
End Function

Private Sub NMDC_ApplyTypedFormatting(ByVal table As ListObject)
    On Error GoTo Handler

    Dim column As ListColumn
    Dim cell As Range
    Dim parsedDate As Date
    Dim headerName As String

    If table.DataBodyRange Is Nothing Then Exit Sub

    For Each column In table.ListColumns
        headerName = CStr(column.Name)
        If NMDC_IsDateHeader(headerName) Then
            For Each cell In column.DataBodyRange.Cells
                If Len(Trim$(CStr(cell.Value))) > 0 Then
                    If NMDC_TryParseDate(cell.Value, parsedDate) Then cell.Value = parsedDate
                End If
            Next cell
            If NMDC_IsDateTimeHeader(headerName) Then
                column.DataBodyRange.NumberFormat = "dd-mmm-yyyy hh:mm"
            Else
                column.DataBodyRange.NumberFormat = "dd-mmm-yyyy"
            End If
        ElseIf NMDC_IsNumericHeader(headerName) Then
            For Each cell In column.DataBodyRange.Cells
                If Len(Trim$(CStr(cell.Value))) > 0 And IsNumeric(cell.Value) Then cell.Value = CDbl(cell.Value)
            Next cell
            column.DataBodyRange.NumberFormat = "#,##0"
        Else
            column.DataBodyRange.NumberFormat = "@"
        End If
    Next column
    Exit Sub
Handler:
    NMDC_LogError "TABLE_FORMAT_ERROR", _
        "Excel loaded the data but could not apply all field formats.", _
        table.Name & " | " & Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_IsDateHeader(ByVal headerName As String) As Boolean
    Select Case UCase$(Trim$(headerName))
        Case "DATE/TIME", "EVENT DATE", "LATEST EVENT DATE", "LAST SUCCESSFUL UPDATE", "CREATED AT", "UPDATED AT"
            NMDC_IsDateHeader = True
        Case Else
            NMDC_IsDateHeader = False
    End Select
End Function

Private Function NMDC_IsDateTimeHeader(ByVal headerName As String) As Boolean
    Select Case UCase$(Trim$(headerName))
        Case "DATE/TIME", "LAST SUCCESSFUL UPDATE", "CREATED AT", "UPDATED AT"
            NMDC_IsDateTimeHeader = True
        Case Else
            NMDC_IsDateTimeHeader = False
    End Select
End Function

Private Function NMDC_IsNumericHeader(ByVal headerName As String) As Boolean
    Select Case UCase$(Trim$(headerName))
        Case "SOURCE ROW", "NEW SOURCES", "CHANGED SOURCES", "REMOVED SOURCES", _
             "ADDED RECORDS", "MODIFIED RECORDS", "REMOVED RECORDS", _
             "REVIEW FLAGS", "CONFLICT FLAGS", "APPROVED DOCUMENTS", _
             "APPROVED REVISIONS", "APPROVED TRANSACTIONS", "PENDING ADDED", _
             "PENDING MODIFIED", "PENDING REMOVED", "PENDING UNCHANGED"
            NMDC_IsNumericHeader = True
        Case Else
            NMDC_IsNumericHeader = False
    End Select
End Function

Private Function NMDC_TryParseDate(ByVal rawValue As Variant, ByRef parsedDate As Date) As Boolean
    On Error GoTo Fallback

    Dim text As String
    Dim datePart As Date
    Dim timePart As Date

    If IsDate(rawValue) And VarType(rawValue) <> vbString Then
        parsedDate = CDate(rawValue)
        NMDC_TryParseDate = True
        Exit Function
    End If

    text = Trim$(CStr(rawValue))
    If Len(text) >= 10 And Mid$(text, 5, 1) = "-" And Mid$(text, 8, 1) = "-" Then
        datePart = DateSerial(CInt(Left$(text, 4)), CInt(Mid$(text, 6, 2)), CInt(Mid$(text, 9, 2)))
        parsedDate = datePart
        If Len(text) >= 16 Then
            timePart = TimeSerial(CInt(Mid$(text, 12, 2)), CInt(Mid$(text, 15, 2)), IIf(Len(text) >= 19, CInt(Mid$(text, 18, 2)), 0))
            parsedDate = datePart + timePart
        End If
        NMDC_TryParseDate = True
        Exit Function
    End If

Fallback:
    On Error GoTo Failed
    If IsDate(rawValue) Then
        parsedDate = CDate(rawValue)
        NMDC_TryParseDate = True
        Exit Function
    End If
Failed:
    NMDC_TryParseDate = False
End Function

Private Sub NMDC_ActivateDocumentLinks(ByVal table As ListObject)
    On Error GoTo Handler

    Dim column As ListColumn
    Dim cell As Range
    Dim target As String

    Set column = Nothing
    On Error Resume Next
    Set column = table.ListColumns("Document Link")
    On Error GoTo Handler
    If column Is Nothing Then Exit Sub
    If column.DataBodyRange Is Nothing Then Exit Sub

    For Each cell In column.DataBodyRange.Cells
        target = Trim$(CStr(cell.Value))
        If Len(target) > 0 Then
            On Error Resume Next
            cell.Hyperlinks.Delete
            On Error GoTo Handler
            If Left$(target, 1) = "#" Then
                table.Parent.Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:=Mid$(target, 2), _
                    TextToDisplay:="Open document", ScreenTip:="Open the source document"
            Else
                table.Parent.Hyperlinks.Add Anchor:=cell, Address:=target, _
                    TextToDisplay:="Open document", ScreenTip:="Open the source document"
            End If
        End If
    Next cell
    Exit Sub
Handler:
    NMDC_LogError "HYPERLINK_REFRESH_ERROR", _
        "Excel loaded the index, but one or more document links could not be activated.", _
        table.Name & " | " & Err.Number & " - " & Err.Description
End Sub

Public Function NMDC_LoadDashboard(ByVal csvPath As String) As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim temp As Worksheet
    Dim imported As Range

    If Not NMDC_FileExists(csvPath) Then
        NMDC_LogError "CSV_MISSING", "Excel could not find the exported dashboard data.", "File: " & csvPath
        NMDC_LoadDashboard = False
        Exit Function
    End If

    Application.ScreenUpdating = False
    If Not NMDC_ImportCsvToTemporarySheet(csvPath, temp, imported) Then GoTo ImportFailed
    Set ws = ThisWorkbook.Worksheets("Home")

    ws.Range("B6").Value = NMDC_ImportedValue(imported, "Approved Status")
    ws.Range("B7").Value = NMDC_ImportedValue(imported, "Approved Run ID")
    ws.Range("B8").Value = NMDC_ImportedValue(imported, "Current Data Folder")
    ws.Range("B9").Value = NMDC_ImportedValue(imported, "Last Successful Update")
    ws.Range("E6").Value = NMDC_ImportedValue(imported, "Approved Documents")
    ws.Range("E7").Value = NMDC_ImportedValue(imported, "Approved Revisions")
    ws.Range("E8").Value = NMDC_ImportedValue(imported, "Approved Transactions")
    ws.Range("H6").Value = NMDC_ImportedValue(imported, "Pending Status")
    ws.Range("H7").Value = NMDC_ImportedValue(imported, "Pending Run ID")
    ws.Range("H8").Value = "Added " & NMDC_ImportedValue(imported, "Pending Added") & _
                           " | Modified " & NMDC_ImportedValue(imported, "Pending Modified") & _
                           " | Removed " & NMDC_ImportedValue(imported, "Pending Removed") & _
                           " | Unchanged " & NMDC_ImportedValue(imported, "Pending Unchanged")
    ws.Range("K6").Value = NMDC_ImportedValue(imported, "Review Flags")
    ws.Range("K7").Value = NMDC_ImportedValue(imported, "Conflict Flags")

    ws.Range("B9").NumberFormat = "dd-mmm-yyyy hh:mm"
    ws.Range("E6:E8").NumberFormat = "#,##0"
    ws.Range("K6:K7").NumberFormat = "#,##0"

    NMDC_DeleteTemporarySheet temp
    Application.ScreenUpdating = True
    NMDC_LoadDashboard = True
    Exit Function

ImportFailed:
    NMDC_DeleteTemporarySheet temp
    Application.ScreenUpdating = True
    NMDC_LoadDashboard = False
    Exit Function

Handler:
    NMDC_DeleteTemporarySheet temp
    Application.ScreenUpdating = True
    NMDC_LogError "DASHBOARD_REFRESH_ERROR", _
        "Excel could not refresh the Home dashboard.", _
        Err.Number & " - " & Err.Description
    NMDC_LoadDashboard = False
End Function

Private Function NMDC_ImportedValue(ByVal imported As Range, ByVal headerName As String) As Variant
    Dim columnIndex As Long
    columnIndex = NMDC_ImportedColumn(imported, headerName)
    If columnIndex = 0 Or imported.Rows.Count < 2 Then
        NMDC_ImportedValue = ""
    Else
        NMDC_ImportedValue = imported.Cells(2, columnIndex).Value
    End If
End Function

Private Function NMDC_ImportedColumn(ByVal imported As Range, ByVal headerName As String) As Long
    Dim index As Long
    For index = 1 To imported.Columns.Count
        If StrComp(Trim$(CStr(imported.Cells(1, index).Value)), headerName, vbTextCompare) = 0 Then
            NMDC_ImportedColumn = index
            Exit Function
        End If
    Next index
    NMDC_ImportedColumn = 0
End Function
