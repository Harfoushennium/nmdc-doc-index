Attribute VB_Name = "modNMDC_Refresh"
Option Explicit

Public Function NMDC_RefreshExchangeData() As Boolean
    On Error GoTo Handler

    Dim refreshOk As Boolean
    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean

    refreshOk = True
    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    NMDC_SetRefreshStatus 1, 8, "Master Documents"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\master_documents.csv", "Master Documents", "MasterDocuments") Then refreshOk = False

    NMDC_SetRefreshStatus 2, 8, "Revisions"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\revisions.csv", "Revisions", "RevisionRegister") Then refreshOk = False

    NMDC_SetRefreshStatus 3, 8, "Transactions"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\events.csv", "Transactions", "EventRegister") Then refreshOk = False

    NMDC_SetRefreshStatus 4, 8, "Pending Update"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\pending_update.csv", "Pending Update", "PendingUpdate") Then refreshOk = False

    NMDC_SetRefreshStatus 5, 8, "Review Flags"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\flags.csv", "Review Flags", "ReviewFlags") Then refreshOk = False

    NMDC_SetRefreshStatus 6, 8, "Update History"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\history.csv", "Update History", "UpdateHistory") Then refreshOk = False

    NMDC_SetRefreshStatus 7, 8, "Error Log"
    If Not NMDC_LoadErrorCsvPreserveLocal(NMDC_ExchangePath() & "\errors.csv") Then refreshOk = False

    NMDC_SetRefreshStatus 8, 8, "Dashboard"
    If Not NMDC_LoadDashboard(NMDC_ExchangePath() & "\dashboard.csv") Then refreshOk = False

    NMDC_ApplyWorkbookGuidance

CleanExit:
    Application.Calculation = previousCalculation
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Application.StatusBar = False
    NMDC_RefreshExchangeData = refreshOk
    Exit Function

Handler:
    refreshOk = False
    NMDC_LogError "EXCEL_REFRESH_ERROR", _
        "Excel could not refresh one or more NMDC Index tables.", _
        Err.Number & " - " & Err.Description
    Resume CleanExit
End Function

Private Sub NMDC_SetRefreshStatus(ByVal currentStep As Long, ByVal totalSteps As Long, ByVal labelText As String)
    Dim filled As Long
    Dim bar As String
    filled = Int((currentStep / totalSteps) * 20)
    bar = "[" & String$(filled, "=") & String$(20 - filled, ".") & "]"
    Application.StatusBar = "NMDC Document Index - refreshing " & labelText & " " & bar & _
                            "  " & CStr(currentStep) & "/" & CStr(totalSteps)
    DoEvents
End Sub

Private Sub NMDC_ResetTableViewForRefresh(ByVal table As ListObject)
    On Error Resume Next

    ' Never resize or replace rows while a ListObject is filtered or sorted.
    ' Hidden/filter state can otherwise make the refreshed display appear to
    ' mix old and new rows even though the underlying array assignment is valid.
    If table.Parent.FilterMode Then table.Parent.ShowAllData
    If table.Sort.SortFields.Count > 0 Then table.Sort.SortFields.Clear
    table.ShowAutoFilter = True

    On Error GoTo 0
End Sub

Public Function NMDC_LoadCsvToTable(ByVal csvPath As String, ByVal sheetName As String, ByVal tableName As String) As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim headers As Variant
    Dim data As Variant
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

    If Not NMDC_ParseCsvFile(csvPath, headers, data, dataCount, columnCount) Then GoTo ImportFailed

    tableRows = dataCount
    If tableRows < 1 Then tableRows = 1
    headerRow = table.HeaderRowRange.Row
    firstColumn = table.Range.Column

    NMDC_ResetTableViewForRefresh table

    If Not table.DataBodyRange Is Nothing Then
        On Error Resume Next
        table.DataBodyRange.Hyperlinks.Delete
        table.DataBodyRange.ClearContents
        On Error GoTo Handler
    End If

    Set targetRange = ws.Range(ws.Cells(headerRow, firstColumn), _
                               ws.Cells(headerRow + tableRows, firstColumn + columnCount - 1))
    table.Resize targetRange
    table.HeaderRowRange.Value2 = headers

    If dataCount > 0 Then
        If table.DataBodyRange Is Nothing Then table.ListRows.Add
        If table.DataBodyRange Is Nothing Then
            Err.Raise vbObjectError + 321, "NMDC CSV Refresh", _
                "Excel could not create a data row for table " & tableName & "."
        End If
        table.DataBodyRange.Value2 = data
    Else
        If table.ListRows.Count = 0 Then table.ListRows.Add
        If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.ClearContents
    End If

    NMDC_ApplyTypedFormatting table
    NMDC_ActivateDocumentLinks table
    NMDC_ActivateSourceLinks table
    NMDC_EnhanceReviewFlagGuidance table
    NMDC_ApplyReviewFlagValidation table
    NMDC_ApplyTableGuidance table

    NMDC_LoadCsvToTable = True
    Exit Function

ImportFailed:
    NMDC_LoadCsvToTable = False
    Exit Function

Handler:
    NMDC_LogError "CSV_IMPORT_ERROR", _
        "Excel could not load " & sheetName & ".", _
        "Table: " & tableName & " | File: " & csvPath & " | " & Err.Number & " - " & Err.Description
    NMDC_LoadCsvToTable = False
End Function

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
    NMDC_ActivateSourceLinks table
    NMDC_ApplyTableGuidance table

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

Private Sub NMDC_ApplyTypedFormatting(ByVal table As ListObject)
    On Error GoTo Handler

    Dim column As ListColumn
    Dim headerName As String

    If table.DataBodyRange Is Nothing Then Exit Sub

    For Each column In table.ListColumns
        headerName = CStr(column.Name)
        If NMDC_IsDateHeader(headerName) Then
            NMDC_ConvertDateColumn column, NMDC_IsDateTimeHeader(headerName)
        ElseIf NMDC_IsNumericHeader(headerName) Then
            NMDC_ConvertNumericColumn column
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

Private Sub NMDC_ConvertDateColumn(ByVal column As ListColumn, ByVal includeTime As Boolean)
    Dim values As Variant
    Dim outputValues() As Variant
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim rawValue As Variant
    Dim parsedDate As Date

    If column.DataBodyRange Is Nothing Then Exit Sub
    rowCount = column.DataBodyRange.Rows.Count
    values = column.DataBodyRange.Value2
    ReDim outputValues(1 To rowCount, 1 To 1)

    For rowIndex = 1 To rowCount
        If rowCount = 1 Then
            rawValue = values
        Else
            rawValue = values(rowIndex, 1)
        End If
        If Len(Trim$(CStr(rawValue))) > 0 And NMDC_TryParseDate(rawValue, parsedDate) Then
            outputValues(rowIndex, 1) = parsedDate
        Else
            outputValues(rowIndex, 1) = rawValue
        End If
    Next rowIndex

    column.DataBodyRange.Value2 = outputValues
    If includeTime Then
        column.DataBodyRange.NumberFormat = "dd-mmm-yyyy hh:mm"
    Else
        column.DataBodyRange.NumberFormat = "dd-mmm-yyyy"
    End If
End Sub

Private Sub NMDC_ConvertNumericColumn(ByVal column As ListColumn)
    Dim values As Variant
    Dim outputValues() As Variant
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim rawValue As Variant

    If column.DataBodyRange Is Nothing Then Exit Sub
    rowCount = column.DataBodyRange.Rows.Count
    values = column.DataBodyRange.Value2
    ReDim outputValues(1 To rowCount, 1 To 1)

    For rowIndex = 1 To rowCount
        If rowCount = 1 Then
            rawValue = values
        Else
            rawValue = values(rowIndex, 1)
        End If
        If Len(Trim$(CStr(rawValue))) > 0 And IsNumeric(rawValue) Then
            outputValues(rowIndex, 1) = CDbl(rawValue)
        Else
            outputValues(rowIndex, 1) = rawValue
        End If
    Next rowIndex

    column.DataBodyRange.Value2 = outputValues
    column.DataBodyRange.NumberFormat = "#,##0"
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
    Dim values As Variant
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim target As String
    Dim targetCell As Range
    Dim previousAutoFill As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    Set column = Nothing
    On Error Resume Next
    Set column = table.ListColumns("Document Link")
    On Error GoTo Handler
    If column Is Nothing Then Exit Sub
    If column.DataBodyRange Is Nothing Then Exit Sub

    rowCount = column.DataBodyRange.Rows.Count
    values = column.DataBodyRange.Value2
    previousAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    Application.AutoCorrect.AutoFillFormulasInLists = False

    On Error Resume Next
    column.DataBodyRange.Hyperlinks.Delete
    On Error GoTo Handler
    column.DataBodyRange.NumberFormat = "General"

    For rowIndex = 1 To rowCount
        If rowCount = 1 Then
            target = Trim$(CStr(values))
        Else
            target = Trim$(CStr(values(rowIndex, 1)))
        End If

        Set targetCell = column.DataBodyRange.Cells(rowIndex, 1)
        targetCell.ClearContents
        If Len(target) > 0 Then
            targetCell.Value2 = "Open document"
            table.Parent.Hyperlinks.Add Anchor:=targetCell, Address:=target, TextToDisplay:="Open document"
        End If
    Next rowIndex

    Application.AutoCorrect.AutoFillFormulasInLists = previousAutoFill
    Exit Sub
Handler:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    Application.AutoCorrect.AutoFillFormulasInLists = previousAutoFill
    On Error GoTo 0
    NMDC_LogError "HYPERLINK_REFRESH_ERROR", _
        "Excel loaded the index, but one or more document links could not be activated.", _
        table.Name & " | " & errNumber & " - " & errDescription
End Sub

Private Sub NMDC_ActivateSourceLinks(ByVal table As ListObject)
    On Error GoTo Handler

    NMDC_ActivateSourceColumn table, "Source File"
    NMDC_ActivateSourceColumn table, "relative_path"
    NMDC_ActivateSourceColumn table, "Relative Path"
    Exit Sub
Handler:
    NMDC_LogError "SOURCE_HYPERLINK_ERROR", _
        "Excel loaded the table but could not activate all source-file links.", _
        table.Name & " | " & Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_ActivateSourceColumn(ByVal table As ListObject, ByVal columnName As String)
    On Error GoTo Handler

    Dim column As ListColumn
    Dim values As Variant
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim sourceText As String
    Dim addressText As String
    Dim dataFolder As String
    Dim fso As Object
    Dim targetCell As Range
    Dim previousAutoFill As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    Set column = Nothing
    On Error Resume Next
    Set column = table.ListColumns(columnName)
    On Error GoTo Handler
    If column Is Nothing Then Exit Sub
    If column.DataBodyRange Is Nothing Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    dataFolder = Trim$(NMDC_ConfigValue("Data Folder"))
    rowCount = column.DataBodyRange.Rows.Count
    values = column.DataBodyRange.Value2
    previousAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    Application.AutoCorrect.AutoFillFormulasInLists = False

    On Error Resume Next
    column.DataBodyRange.Hyperlinks.Delete
    On Error GoTo Handler
    column.DataBodyRange.NumberFormat = "@"

    For rowIndex = 1 To rowCount
        If rowCount = 1 Then
            sourceText = Trim$(CStr(values))
        Else
            sourceText = Trim$(CStr(values(rowIndex, 1)))
        End If

        Set targetCell = column.DataBodyRange.Cells(rowIndex, 1)
        targetCell.ClearContents
        targetCell.Value2 = sourceText

        If Len(sourceText) > 0 Then
            addressText = sourceText
            If Len(dataFolder) > 0 Then
                If Len(fso.GetDriveName(sourceText)) = 0 And Left$(sourceText, 2) <> "\\" Then
                    addressText = fso.BuildPath(dataFolder, Replace(sourceText, "/", "\"))
                End If
            End If
            table.Parent.Hyperlinks.Add Anchor:=targetCell, Address:=addressText, TextToDisplay:=sourceText
        End If
    Next rowIndex

    Application.AutoCorrect.AutoFillFormulasInLists = previousAutoFill
    Exit Sub
Handler:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    Application.AutoCorrect.AutoFillFormulasInLists = previousAutoFill
    On Error GoTo 0
    If errNumber = 0 Then errNumber = vbObjectError + 322
    Err.Raise errNumber, "NMDC Source Hyperlink Refresh", errDescription
End Sub

Private Sub NMDC_EnhanceReviewFlagGuidance(ByVal table As ListObject)
    On Error GoTo Handler
    If StrComp(table.Name, "ReviewFlags", vbTextCompare) <> 0 Then Exit Sub
    If table.DataBodyRange Is Nothing Then Exit Sub

    Dim row As ListRow
    Dim codeCol As Long
    Dim problemCol As Long
    Dim actionCol As Long
    Dim flagCode As String

    codeCol = table.ListColumns("Flag Code").Index
    problemCol = table.ListColumns("Plain-English Problem").Index
    actionCol = table.ListColumns("Recommended User Action").Index

    For Each row In table.ListRows
        flagCode = UCase$(Trim$(CStr(row.Range.Cells(1, codeCol).Value)))
        If flagCode = "UNRECOGNIZED_LAYOUT" Then
            row.Range.Cells(1, problemCol).Value = _
                "The parser could not identify a safe document-number/data-row layout automatically. " & _
                "This does not necessarily mean the workbook is badly formatted."
            row.Range.Cells(1, actionCol).Value = _
                "Open the Source File and named Source Sheet. If it is empty, choose NO ACTION REQUIRED. " & _
                "If it contains normal register data, tick Select? for the affected row(s), then click Report Selected Parser Fix. " & _
                "The workbook writes the fix report beside this Excel file and opens its location. " & _
                "After a corrected parser/config is installed, click Retry After Fix."
        End If
    Next row
    Exit Sub
Handler:
    NMDC_LogError "REVIEW_GUIDANCE_ERROR", _
        "Excel loaded Review Flags but could not apply the user guidance.", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_ApplyReviewFlagValidation(ByVal table As ListObject)
    On Error GoTo Handler
    If StrComp(table.Name, "ReviewFlags", vbTextCompare) <> 0 Then Exit Sub
    If table.DataBodyRange Is Nothing Then Exit Sub

    Dim selectRange As Range
    Dim decisionRange As Range
    Dim statusRange As Range
    Dim commentRange As Range
    Dim cell As Range

    Set selectRange = table.ListColumns("Select?").DataBodyRange
    Set decisionRange = table.ListColumns("User Decision").DataBodyRange
    Set commentRange = table.ListColumns("User Comment").DataBodyRange
    Set statusRange = table.ListColumns("Resolution Status").DataBodyRange

    For Each cell In selectRange.Cells
        If Len(Trim$(CStr(cell.Value))) = 0 Then cell.Value = False
    Next cell
    On Error Resume Next
    selectRange.CellControl.SetCheckbox
    If Err.Number <> 0 Then
        NMDC_LogError "REVIEW_CHECKBOX_UNAVAILABLE", _
            "Excel could not display Review Flags selection checkboxes. TRUE/FALSE selection values remain usable.", _
            Err.Number & " - " & Err.Description
        Err.Clear
    Else
        selectRange.HorizontalAlignment = xlCenter
    End If
    On Error GoTo Handler

    decisionRange.Validation.Delete
    decisionRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:="ACKNOWLEDGED,NO ACTION REQUIRED,NEEDS SOURCE CORRECTION,NEEDS PARSER/MAPPING FIX,HOLD FOR REVIEW"
    decisionRange.Validation.IgnoreBlank = True
    decisionRange.Validation.InCellDropdown = True
    decisionRange.Validation.ShowInput = True
    decisionRange.Validation.InputTitle = "What should happen?"
    decisionRange.Validation.InputMessage = "Select a review decision. The decision records your review; it does not directly rewrite source/extracted data. See the header note for each choice."

    statusRange.Validation.Delete
    statusRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:="OPEN,ACKNOWLEDGED,RESOLVED,DEFERRED"
    statusRange.Validation.IgnoreBlank = True
    statusRange.Validation.InCellDropdown = True
    statusRange.Validation.ShowInput = True
    statusRange.Validation.InputTitle = "Review status"
    statusRange.Validation.InputMessage = "OPEN = unresolved; ACKNOWLEDGED = reviewed; RESOLVED = closed; DEFERRED = postponed."

    selectRange.Interior.Pattern = xlNone
    decisionRange.Interior.Pattern = xlNone
    commentRange.Interior.Pattern = xlNone
    statusRange.Interior.Pattern = xlNone
    decisionRange.Font.ColorIndex = xlAutomatic
    commentRange.Font.ColorIndex = xlAutomatic
    statusRange.Font.ColorIndex = xlAutomatic
    Exit Sub
Handler:
    NMDC_LogError "REVIEW_DROPDOWN_ERROR", _
        "Excel loaded Review Flags but could not apply one or more dropdown menus.", _
        Err.Number & " - " & Err.Description
End Sub

Public Function NMDC_LoadDashboard(ByVal csvPath As String) As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim headers As Variant
    Dim data As Variant
    Dim dataCount As Long
    Dim columnCount As Long
    Dim parsedDate As Date

    If Not NMDC_FileExists(csvPath) Then
        NMDC_LogError "CSV_MISSING", "Excel could not find the exported dashboard data.", "File: " & csvPath
        NMDC_LoadDashboard = False
        Exit Function
    End If

    If Not NMDC_ParseCsvFile(csvPath, headers, data, dataCount, columnCount) Then
        NMDC_LoadDashboard = False
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets("Home")
    ws.Range("B6").Value = NMDC_CsvValue(headers, data, dataCount, "Approved Status")
    ws.Range("B7").Value = NMDC_CsvValue(headers, data, dataCount, "Approved Run ID")
    ws.Range("B8").Value = NMDC_CsvValue(headers, data, dataCount, "Current Data Folder")

    If NMDC_TryParseDate(NMDC_CsvValue(headers, data, dataCount, "Last Successful Update"), parsedDate) Then
        ws.Range("B9").Value = parsedDate
    Else
        ws.Range("B9").Value = NMDC_CsvValue(headers, data, dataCount, "Last Successful Update")
    End If

    ws.Range("E6").Value = Val(CStr(NMDC_CsvValue(headers, data, dataCount, "Approved Documents")))
    ws.Range("E7").Value = Val(CStr(NMDC_CsvValue(headers, data, dataCount, "Approved Revisions")))
    ws.Range("E8").Value = Val(CStr(NMDC_CsvValue(headers, data, dataCount, "Approved Transactions")))
    ws.Range("H6").Value = NMDC_CsvValue(headers, data, dataCount, "Pending Status")
    ws.Range("H7").Value = NMDC_CsvValue(headers, data, dataCount, "Pending Run ID")
    ws.Range("H8").Value = "Added " & NMDC_CsvValue(headers, data, dataCount, "Pending Added") & _
                           " | Modified " & NMDC_CsvValue(headers, data, dataCount, "Pending Modified") & _
                           " | Removed " & NMDC_CsvValue(headers, data, dataCount, "Pending Removed") & _
                           " | Unchanged " & NMDC_CsvValue(headers, data, dataCount, "Pending Unchanged")
    ws.Range("K6").Value = Val(CStr(NMDC_CsvValue(headers, data, dataCount, "Review Flags")))
    ws.Range("K7").Value = Val(CStr(NMDC_CsvValue(headers, data, dataCount, "Conflict Flags")))

    ws.Range("B9").NumberFormat = "dd-mmm-yyyy hh:mm"
    ws.Range("E6:E8").NumberFormat = "#,##0"
    ws.Range("K6:K7").NumberFormat = "#,##0"

    NMDC_LoadDashboard = True
    Exit Function
Handler:
    NMDC_LogError "DASHBOARD_REFRESH_ERROR", _
        "Excel could not refresh the Home dashboard.", _
        Err.Number & " - " & Err.Description
    NMDC_LoadDashboard = False
End Function

Private Function NMDC_CsvValue(ByVal headers As Variant, ByVal data As Variant, ByVal dataCount As Long, ByVal headerName As String) As Variant
    Dim columnIndex As Long
    columnIndex = NMDC_CsvColumn(headers, headerName)
    If columnIndex = 0 Or dataCount < 1 Then
        NMDC_CsvValue = ""
    Else
        NMDC_CsvValue = data(1, columnIndex)
    End If
End Function

Private Function NMDC_CsvColumn(ByVal headers As Variant, ByVal headerName As String) As Long
    Dim index As Long
    For index = 1 To UBound(headers, 2)
        If StrComp(Trim$(CStr(headers(1, index))), headerName, vbTextCompare) = 0 Then
            NMDC_CsvColumn = index
            Exit Function
        End If
    Next index
    NMDC_CsvColumn = 0
End Function
