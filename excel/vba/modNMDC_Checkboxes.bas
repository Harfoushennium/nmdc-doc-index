Attribute VB_Name = "modNMDC_Checkboxes"
Option Explicit

Private Const SOURCE_CHECK_PREFIX As String = "NMDC_SourceCheck_"

Public Sub NMDC_RebuildSourceSelectionCheckboxes()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim sourceColumn As ListColumn
    Dim includeColumn As ListColumn
    Dim rowIndex As Long
    Dim targetCell As Range
    Dim sourceFile As String
    Dim checkBox As Object
    Dim item As Object
    Dim isIncluded As Boolean

    Set ws = ThisWorkbook.Worksheets("Source Selection")
    Set table = ws.ListObjects("SourceSelection")
    Set includeColumn = table.ListColumns("Include in Index?")
    Set sourceColumn = table.ListColumns("Source File")

    ' Delete only NMDC-generated Form Control checkboxes. Do not disturb any
    ' unrelated workbook controls or user content.
    For rowIndex = ws.CheckBoxes.Count To 1 Step -1
        Set item = ws.CheckBoxes(rowIndex)
        If Left$(CStr(item.Name), Len(SOURCE_CHECK_PREFIX)) = SOURCE_CHECK_PREFIX Then item.Delete
    Next rowIndex

    If table.DataBodyRange Is Nothing Then Exit Sub

    includeColumn.Range.ColumnWidth = 12
    includeColumn.Range.HorizontalAlignment = xlCenter

    For rowIndex = 1 To table.ListRows.Count
        sourceFile = Trim$(CStr(sourceColumn.DataBodyRange.Cells(rowIndex, 1).Value))
        Set targetCell = includeColumn.DataBodyRange.Cells(rowIndex, 1)
        If Len(sourceFile) > 0 Then
            isIncluded = NMDC_SourceCheckedValue(targetCell.Value)
            targetCell.Value = isIncluded
            ' Keep the logical TRUE/FALSE value in the cell for audit/export,
            ' but hide the text because the checkbox is the user interface.
            targetCell.NumberFormat = ";;;"

            Set checkBox = ws.CheckBoxes.Add( _
                targetCell.Left + (targetCell.Width - 13) / 2, _
                targetCell.Top + (targetCell.Height - 13) / 2, _
                13, 13)
            checkBox.Name = SOURCE_CHECK_PREFIX & Format$(rowIndex, "0000")
            checkBox.Caption = ""
            checkBox.Value = IIf(isIncluded, xlOn, xlOff)
            checkBox.OnAction = "NMDC_SourceCheckboxClicked"
            checkBox.Placement = xlMoveAndSize
            checkBox.PrintObject = False
        Else
            targetCell.ClearContents
        End If
    Next rowIndex
    Exit Sub

Handler:
    NMDC_LogError "SOURCE_CHECKBOX_BUILD_ERROR", _
        "Excel could not build the Source Selection checkboxes.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_SourceCheckboxClicked()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim checkBox As Object
    Dim rowIndex As Long
    Dim targetCell As Range

    Set ws = ThisWorkbook.Worksheets("Source Selection")
    Set table = ws.ListObjects("SourceSelection")
    Set checkBox = ws.CheckBoxes(CStr(Application.Caller))
    rowIndex = checkBox.TopLeftCell.Row - table.DataBodyRange.Row + 1
    If rowIndex < 1 Or rowIndex > table.ListRows.Count Then Exit Sub

    Set targetCell = table.ListColumns("Include in Index?").DataBodyRange.Cells(rowIndex, 1)
    targetCell.Value = (checkBox.Value = xlOn)
    targetCell.NumberFormat = ";;;"
    ws.Range("A3").Value = "Selection changed - click Save Selection & Restage when finished."
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Color = RGB(156, 104, 0)
    Exit Sub

Handler:
    NMDC_LogError "SOURCE_CHECKBOX_CLICK_ERROR", _
        "Excel could not record the checkbox selection.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_CheckAllSources()
    NMDC_SetAllSourceCheckboxes True
End Sub

Public Sub NMDC_UncheckAllSources()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Uncheck every source workbook?" & vbCrLf & vbCrLf & _
                    "Nothing is approved automatically. You must still click Save Selection & Restage.", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer = vbYes Then NMDC_SetAllSourceCheckboxes False
End Sub

Private Sub NMDC_SetAllSourceCheckboxes(ByVal checkedValue As Boolean)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim targetCell As Range
    Dim checkBox As Object

    Set ws = ThisWorkbook.Worksheets("Source Selection")
    Set table = ws.ListObjects("SourceSelection")
    If table.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To table.ListRows.Count
        sourceFile = Trim$(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value))
        If Len(sourceFile) > 0 Then
            Set targetCell = table.ListColumns("Include in Index?").DataBodyRange.Cells(rowIndex, 1)
            targetCell.Value = checkedValue
            targetCell.NumberFormat = ";;;"
            Set checkBox = Nothing
            On Error Resume Next
            Set checkBox = ws.CheckBoxes(SOURCE_CHECK_PREFIX & Format$(rowIndex, "0000"))
            On Error GoTo Handler
            If Not checkBox Is Nothing Then checkBox.Value = IIf(checkedValue, xlOn, xlOff)
        End If
    Next rowIndex

    ws.Range("A3").Value = "Selection changed - click Save Selection & Restage when finished."
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Color = RGB(156, 104, 0)
    Exit Sub

Handler:
    NMDC_LogError "SOURCE_CHECKBOX_ALL_ERROR", _
        "Excel could not update all source checkboxes.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_SaveSourceSelections()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim ownerNote As String
    Dim includeValue As Boolean
    Dim csvText As String
    Dim decisionsPath As String
    Dim args As String
    Dim answer As VbMsgBoxResult

    Set ws = ThisWorkbook.Worksheets("Source Selection")
    Set table = ws.ListObjects("SourceSelection")

    answer = MsgBox("Save the checked/unchecked source selection and restage the proposal?" & vbCrLf & vbCrLf & _
                    "Checked = include in index scope" & vbCrLf & _
                    "Unchecked = intentionally exclude this source workbook" & vbCrLf & vbCrLf & _
                    "The approved index will NOT change until you later press Approve Update.", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer <> vbYes Then Exit Sub

    csvText = NMDC_CheckboxCsvField("Source File") & "," & _
              NMDC_CheckboxCsvField("Include in Index?") & "," & _
              NMDC_CheckboxCsvField("Owner Note") & vbLf

    If Not table.DataBodyRange Is Nothing Then
        For rowIndex = 1 To table.ListRows.Count
            sourceFile = Trim$(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value))
            If Len(sourceFile) > 0 Then
                includeValue = NMDC_SourceCheckedValue(table.ListColumns("Include in Index?").DataBodyRange.Cells(rowIndex, 1).Value)
                ownerNote = Trim$(CStr(table.ListColumns("Owner Note").DataBodyRange.Cells(rowIndex, 1).Value))
                csvText = csvText & NMDC_CheckboxCsvField(sourceFile) & "," & _
                          NMDC_CheckboxCsvField(IIf(includeValue, "TRUE", "FALSE")) & "," & _
                          NMDC_CheckboxCsvField(ownerNote) & vbLf
            End If
        Next rowIndex
    End If

    decisionsPath = NMDC_ExchangePath() & "\source_selection_decisions.csv"
    NMDC_WriteCheckboxUtf8 decisionsPath, csvText
    args = "--selections-file " & NMDC_Quote(decisionsPath)

    If NMDC_RunEngine("save-source-selections", args) <> 0 Then
        MsgBox "The source selection could not be saved. Please review Error Log.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    ws.Range("A3").Value = "Source selection saved. Restaging proposal..."
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Color = RGB(31, 70, 90)
    NMDC_UpdateChangedFilesFast
    Exit Sub

Handler:
    NMDC_LogError "SOURCE_CHECKBOX_SAVE_ERROR", _
        "Excel could not save Source Selection checkbox choices.", _
        Err.Number & " - " & Err.Description
    MsgBox "Source Selection could not be saved. Please review Error Log.", _
           vbExclamation, "NMDC Document Index"
End Sub

Private Function NMDC_SourceCheckedValue(ByVal value As Variant) As Boolean
    If VarType(value) = vbBoolean Then
        NMDC_SourceCheckedValue = CBool(value)
        Exit Function
    End If
    Select Case UCase$(Trim$(CStr(value)))
        Case "TRUE", "YES", "1", "ON", "CHECKED", "INCLUDE", "INCLUDED"
            NMDC_SourceCheckedValue = True
        Case Else
            NMDC_SourceCheckedValue = False
    End Select
End Function

Private Function NMDC_CheckboxCsvField(ByVal value As String) As String
    NMDC_CheckboxCsvField = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Sub NMDC_WriteCheckboxUtf8(ByVal filePath As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.SaveToFile filePath, 2
    stream.Close
End Sub
