Attribute VB_Name = "modNMDC_TableActions"
Option Explicit

Public Sub NMDC_FlagWrongDataFromTable()
    On Error GoTo Handler

    Dim selectedCell As Range
    Dim table As ListObject
    Dim dataRow As ListRow
    Dim projectNo As String
    Dim documentNo As String
    Dim sourceFile As String
    Dim sourceSheet As String
    Dim sourceRow As String
    Dim sourceCell As String
    Dim revision As String
    Dim eventIdentity As String
    Dim flagCode As String
    Dim currentValue As String
    Dim currentField As String
    Dim userNote As String
    Dim expectedValue As String
    Dim extraArgs As String

    Set selectedCell = NMDC_SelectedDataCell()
    If selectedCell Is Nothing Then Exit Sub

    Set table = NMDC_TableForCell(selectedCell)
    If table Is Nothing Then
        MsgBox "The selected cell is not inside a valid NMDC data table." & vbCrLf & vbCrLf & _
               "Please select a data cell inside one of the filtered Excel tables.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    Set dataRow = NMDC_TableRowForCell(table, selectedCell)
    If dataRow Is Nothing Then
        MsgBox "Please select a data row, not the table header.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    projectNo = NMDC_TableRowValue(table, dataRow, "Project No.")
    documentNo = NMDC_TableRowValue(table, dataRow, "Document No.")
    sourceFile = NMDC_TableRowValue(table, dataRow, "Source File")
    sourceSheet = NMDC_TableRowValue(table, dataRow, "Source Sheet")
    sourceRow = NMDC_TableRowValue(table, dataRow, "Source Row")
    sourceCell = NMDC_TableRowValue(table, dataRow, "Source Cell")
    revision = NMDC_TableRowValue(table, dataRow, "Revision")
    eventIdentity = NMDC_TableRowValue(table, dataRow, "Event Key")
    If Len(eventIdentity) = 0 Then eventIdentity = NMDC_TableRowValue(table, dataRow, "Record Identity")
    flagCode = NMDC_TableRowValue(table, dataRow, "Flag Code")
    currentField = NMDC_TableHeaderForCell(table, selectedCell)
    currentValue = NMDC_SelectedTableCellValue(selectedCell)

    userNote = InputBox("Describe what is wrong with this record:", "Flag Wrong Data")
    If Len(Trim$(userNote)) = 0 Then Exit Sub
    expectedValue = InputBox("Enter the correct/expected value if known (optional):", "Flag Wrong Data")

    extraArgs = "--message " & NMDC_Quote(userNote) & _
                " --project-no " & NMDC_Quote(projectNo) & _
                " --document-no " & NMDC_Quote(documentNo) & _
                " --source-file " & NMDC_Quote(sourceFile) & _
                " --worksheet " & NMDC_Quote(sourceSheet) & _
                " --source-row " & NMDC_Quote(sourceRow) & _
                " --source-cell " & NMDC_Quote(sourceCell) & _
                " --revision " & NMDC_Quote(revision) & _
                " --event-identity " & NMDC_Quote(eventIdentity) & _
                " --flag-code " & NMDC_Quote(flagCode) & _
                " --current-field " & NMDC_Quote(currentField) & _
                " --current-value " & NMDC_Quote(currentValue) & _
                " --expected-value " & NMDC_Quote(expectedValue) & _
                " --user-name " & NMDC_Quote(Application.UserName)

    If NMDC_RunEngine("user-flag", extraArgs) = 0 Then
        MsgBox "Your flag has been recorded for review.", vbInformation, "NMDC Document Index"
    End If
    Exit Sub

Handler:
    NMDC_LogError "USER_FLAG_ERROR", "Excel could not record the user flag.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ReportRequirementFromTable()
    On Error GoTo Handler

    Dim message As String
    Dim contextArgs As String
    Dim selectedCell As Range
    Dim table As ListObject
    Dim dataRow As ListRow

    message = InputBox("Describe the problem or new requirement in plain English:", "Report Requirement / Problem")
    If Len(Trim$(message)) = 0 Then Exit Sub

    contextArgs = " --user-name " & NMDC_Quote(Application.UserName)
    Set selectedCell = NMDC_CurrentDataCell()
    If Not selectedCell Is Nothing Then
        Set table = NMDC_TableForCell(selectedCell)
        If Not table Is Nothing Then
            Set dataRow = NMDC_TableRowForCell(table, selectedCell)
            If Not dataRow Is Nothing Then
                contextArgs = contextArgs & _
                              " --project-no " & NMDC_Quote(NMDC_TableRowValue(table, dataRow, "Project No.")) & _
                              " --document-no " & NMDC_Quote(NMDC_TableRowValue(table, dataRow, "Document No.")) & _
                              " --revision " & NMDC_Quote(NMDC_TableRowValue(table, dataRow, "Revision")) & _
                              " --source-file " & NMDC_Quote(NMDC_TableRowValue(table, dataRow, "Source File")) & _
                              " --worksheet " & NMDC_Quote(NMDC_TableRowValue(table, dataRow, "Source Sheet")) & _
                              " --source-row " & NMDC_Quote(NMDC_TableRowValue(table, dataRow, "Source Row")) & _
                              " --source-cell " & NMDC_Quote(NMDC_TableRowValue(table, dataRow, "Source Cell")) & _
                              " --event-identity " & NMDC_Quote(NMDC_TableRowValue(table, dataRow, "Event Key")) & _
                              " --current-field " & NMDC_Quote(NMDC_TableHeaderForCell(table, selectedCell)) & _
                              " --current-value " & NMDC_Quote(NMDC_SelectedTableCellValue(selectedCell))
            End If
        End If
    End If

    If NMDC_RunEngine("support-request", "--message " & NMDC_Quote(message) & contextArgs) = 0 Then
        MsgBox "A support package was created successfully." & vbCrLf & _
               "You can provide that package to GPT/developer support.", vbInformation, "NMDC Document Index"
    End If
    Exit Sub

Handler:
    NMDC_LogError "SUPPORT_REQUEST_ERROR", "Excel could not create the support request.", Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_SelectedDataCell() As Range
    Dim selectedCell As Range

    Set selectedCell = NMDC_CurrentDataCell()
    If Not selectedCell Is Nothing Then
        Set NMDC_SelectedDataCell = selectedCell
        Exit Function
    End If

    On Error Resume Next
    Set selectedCell = Application.InputBox( _
        Prompt:="Select the incorrect cell inside one of the NMDC data tables, then press OK.", _
        Title:="Flag Wrong Data", Type:=8)
    On Error GoTo 0

    If selectedCell Is Nothing Then
        Set NMDC_SelectedDataCell = Nothing
    Else
        Set NMDC_SelectedDataCell = selectedCell.Cells(1, 1)
    End If
End Function

Private Function NMDC_CurrentDataCell() As Range
    On Error GoTo Missing
    Dim selectedCell As Range
    Set selectedCell = ActiveCell
    If selectedCell Is Nothing Then GoTo Missing
    If NMDC_TableForCell(selectedCell) Is Nothing Then GoTo Missing
    Set NMDC_CurrentDataCell = selectedCell
    Exit Function
Missing:
    Set NMDC_CurrentDataCell = Nothing
End Function

Private Function NMDC_TableForCell(ByVal selectedCell As Range) As ListObject
    On Error GoTo Missing
    Dim table As ListObject
    For Each table In selectedCell.Worksheet.ListObjects
        If Not Intersect(selectedCell, table.Range) Is Nothing Then
            Set NMDC_TableForCell = table
            Exit Function
        End If
    Next table
Missing:
    Set NMDC_TableForCell = Nothing
End Function

Private Function NMDC_TableRowForCell(ByVal table As ListObject, ByVal selectedCell As Range) As ListRow
    On Error GoTo Missing
    Dim relativeRow As Long
    If table.DataBodyRange Is Nothing Then GoTo Missing
    If Intersect(selectedCell, table.DataBodyRange) Is Nothing Then GoTo Missing
    relativeRow = selectedCell.Row - table.DataBodyRange.Row + 1
    Set NMDC_TableRowForCell = table.ListRows(relativeRow)
    Exit Function
Missing:
    Set NMDC_TableRowForCell = Nothing
End Function

Private Function NMDC_TableRowValue(ByVal table As ListObject, ByVal dataRow As ListRow, ByVal headerName As String) As String
    On Error GoTo Missing
    Dim columnIndex As Long
    columnIndex = table.ListColumns(headerName).Index
    NMDC_TableRowValue = CStr(dataRow.Range.Cells(1, columnIndex).Value)
    Exit Function
Missing:
    NMDC_TableRowValue = ""
End Function

Private Function NMDC_TableHeaderForCell(ByVal table As ListObject, ByVal selectedCell As Range) As String
    On Error GoTo Missing
    Dim relativeColumn As Long
    relativeColumn = selectedCell.Column - table.Range.Column + 1
    If relativeColumn < 1 Or relativeColumn > table.ListColumns.Count Then GoTo Missing
    NMDC_TableHeaderForCell = CStr(table.HeaderRowRange.Cells(1, relativeColumn).Value)
    Exit Function
Missing:
    NMDC_TableHeaderForCell = ""
End Function

Private Function NMDC_SelectedTableCellValue(ByVal cell As Range) As String
    On Error GoTo Handler
    If cell.Hyperlinks.Count > 0 Then
        If Len(cell.Hyperlinks(1).Address) > 0 Then
            NMDC_SelectedTableCellValue = cell.Hyperlinks(1).Address
        Else
            NMDC_SelectedTableCellValue = cell.Hyperlinks(1).SubAddress
        End If
    Else
        NMDC_SelectedTableCellValue = CStr(cell.Value)
    End If
    Exit Function
Handler:
    NMDC_SelectedTableCellValue = CStr(cell.Value)
End Function
