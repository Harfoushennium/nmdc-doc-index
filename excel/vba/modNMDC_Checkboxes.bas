Attribute VB_Name = "modNMDC_Checkboxes"
Option Explicit

Private Const PENDING_SOURCE_CHECK_PREFIX As String = "NMDC_PendingSourceCheck_"
Private Const SOURCE_PANEL_TABLE As String = "SourceSelection"
Private mOwnerEnhancementsReady As Boolean

' Source selection is intentionally integrated into Pending Update.
' The old separate Source Selection worksheet is retained only as a hidden
' compatibility shell; the owner no longer needs to navigate to it.
'
' The source checkboxes use LinkedCell and NO OnAction callback. This removes
' the macro-not-available failure that was observed with NMDC_SourceCheckboxClicked.

Public Sub NMDC_RebuildSourceSelectionCheckboxes()
    ' Backward-compatible entry point used by older owner-UX code.
    NMDC_RebuildPendingSourcePanel
End Sub

Public Sub NMDC_RebuildPendingSourcePanel()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject

    NMDC_EnsureOwnerEnhancements
    NMDC_RetireSeparateSourceSelectionUI

    Set ws = ThisWorkbook.Worksheets("Pending Update")
    Set table = NMDC_EnsurePendingSourceTable(ws)
    If table Is Nothing Then Exit Sub

    If NMDC_FileExists(NMDC_ExchangePath() & "\source_selection.csv") Then
        If Not NMDC_LoadSourceSelectionPanel(table, NMDC_ExchangePath() & "\source_selection.csv") Then
            GoTo Handler
        End If
    End If

    NMDC_FormatPendingSourcePanel ws, table
    NMDC_BuildPendingSourceCheckboxes ws, table
    Exit Sub

Handler:
    NMDC_LogError "PENDING_SOURCE_PANEL_ERROR", _
        "Excel could not prepare source selection inside Pending Update.", _
        Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_EnsurePendingSourceTable(ByVal ws As Worksheet) As ListObject
    On Error GoTo Handler

    Dim table As ListObject
    Dim legacyWs As Worksheet
    Dim legacyTable As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim sourceRange As Range

    Set table = Nothing
    On Error Resume Next
    Set table = ws.ListObjects(SOURCE_PANEL_TABLE)
    On Error GoTo Handler

    If table Is Nothing Then
        Set legacyWs = Nothing
        Set legacyTable = Nothing
        On Error Resume Next
        Set legacyWs = ThisWorkbook.Worksheets("Source Selection")
        If Not legacyWs Is Nothing Then Set legacyTable = legacyWs.ListObjects(SOURCE_PANEL_TABLE)
        On Error GoTo Handler
        If Not legacyTable Is Nothing Then legacyTable.Unlist

        headers = Array( _
            "Include in Index?", "Project No.", "Source File", "Source Family", _
            "Current Status", "Owner Note", "Selection Reason", "Last Processed Run")

        On Error Resume Next
        ws.Range("L4:S7").UnMerge
        On Error GoTo Handler
        For i = LBound(headers) To UBound(headers)
            ws.Cells(5, 12 + i - LBound(headers)).Value = headers(i)
        Next i

        Set sourceRange = ws.Range("L5:S6")
        Set table = ws.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
        table.Name = SOURCE_PANEL_TABLE
        table.TableStyle = "TableStyleMedium2"
        If table.ListRows.Count = 0 Then table.ListRows.Add
    End If

    Set NMDC_EnsurePendingSourceTable = table
    Exit Function

Handler:
    NMDC_LogError "PENDING_SOURCE_TABLE_ERROR", _
        "Excel could not create the source-selection table on Pending Update.", _
        Err.Number & " - " & Err.Description
    Set NMDC_EnsurePendingSourceTable = Nothing
End Function

Private Function NMDC_LoadSourceSelectionPanel(ByVal table As ListObject, ByVal csvPath As String) As Boolean
    On Error GoTo Handler

    Dim headers As Variant
    Dim data As Variant
    Dim dataCount As Long
    Dim columnCount As Long
    Dim tableRows As Long
    Dim headerRow As Long
    Dim firstColumn As Long
    Dim targetRange As Range

    If Not NMDC_ParseCsvFile(csvPath, headers, data, dataCount, columnCount) Then Exit Function

    tableRows = dataCount
    If tableRows < 1 Then tableRows = 1
    headerRow = table.HeaderRowRange.Row
    firstColumn = table.Range.Column

    If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.ClearContents

    Set targetRange = table.Parent.Range( _
        table.Parent.Cells(headerRow, firstColumn), _
        table.Parent.Cells(headerRow + tableRows, firstColumn + columnCount - 1))
    table.Resize targetRange
    table.HeaderRowRange.Value2 = headers

    If dataCount > 0 Then
        table.DataBodyRange.Value2 = data
    Else
        If table.ListRows.Count = 0 Then table.ListRows.Add
        table.DataBodyRange.ClearContents
    End If

    NMDC_ApplyTableGuidance table
    NMDC_LoadSourceSelectionPanel = True
    Exit Function

Handler:
    NMDC_LogError "SOURCE_PANEL_REFRESH_ERROR", _
        "Excel could not refresh the source-selection panel.", _
        Err.Number & " - " & Err.Description
    NMDC_LoadSourceSelectionPanel = False
End Function

Private Sub NMDC_FormatPendingSourcePanel(ByVal ws As Worksheet, ByVal table As ListObject)
    On Error GoTo Handler

    Dim area As Range
    Dim button As Shape

    On Error Resume Next
    ws.Range("L4:S4").UnMerge
    ws.Range("L4:S4").Merge
    On Error GoTo Handler
    With ws.Range("L4:S4")
        .Value = "SOURCE SELECTION - CHECKED = INCLUDE / UNCHECKED = EXCLUDE. Project No. is shown beside each source. Save once after making all choices."
        .Interior.Color = RGB(232, 241, 247)
        .Font.Name = "Aptos"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(31, 70, 90)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 42
    End With

    On Error Resume Next
    ws.Shapes("NMDC_Pending_Exclude").Delete
    ws.Shapes("NMDC_Pending_SourceList").Delete
    ws.Shapes("NMDC_Pending_SourceSave").Delete
    ws.Shapes("NMDC_Pending_SourceAll").Delete
    ws.Shapes("NMDC_Pending_SourceNone").Delete
    On Error GoTo Handler

    Set area = ws.Range("L2:N3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_SourceSave"
    button.OnAction = "NMDC_SaveSourceSelections"
    button.TextFrame.Characters.Text = "Save Source Choices" & vbLf & "& Restage"
    NMDC_FormatSourceButton button, RGB(20, 108, 148)

    Set area = ws.Range("O2:P3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_SourceAll"
    button.OnAction = "NMDC_CheckAllSources"
    button.TextFrame.Characters.Text = "Check All"
    NMDC_FormatSourceButton button, RGB(46, 125, 50)

    Set area = ws.Range("Q2:R3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_SourceNone"
    button.OnAction = "NMDC_UncheckAllSources"
    button.TextFrame.Characters.Text = "Uncheck All"
    NMDC_FormatSourceButton button, RGB(198, 40, 40)

    table.ListColumns("Include in Index?").Range.ColumnWidth = 12
    table.ListColumns("Project No.").Range.ColumnWidth = 16
    table.ListColumns("Source File").Range.ColumnWidth = 42
    table.ListColumns("Owner Note").Range.ColumnWidth = 28

    On Error Resume Next
    ws.Range("A4:I4").Value = _
        "REVIEW WORKFLOW. Review staged changes in the left table. Source selection is now on THIS sheet at the right: checked = include source, unchecked = exclude source. Click Save Source Choices & Restage after changes. Approve/Hold/Reject still applies to the whole staged proposal."
    ThisWorkbook.Worksheets("Home").Range("E34").Value = _
        "In Pending Update, use the source checkbox panel on the right. Project No. and Source File identify each source. Save Source Choices & Restage once when finished."
    On Error GoTo 0
    Exit Sub

Handler:
    NMDC_LogError "PENDING_SOURCE_PANEL_FORMAT_ERROR", _
        "Excel could not format the Pending Update source-selection panel.", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_FormatSourceButton(ByVal button As Shape, ByVal fillColor As Long)
    With button
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = 3
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = fillColor
        .TextFrame.Characters.Font.Name = "Aptos"
        .TextFrame.Characters.Font.Size = 8.5
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Placement = xlMoveAndSize
    End With
End Sub

Private Sub NMDC_BuildPendingSourceCheckboxes(ByVal ws As Worksheet, ByVal table As ListObject)
    On Error GoTo Handler

    Dim includeColumn As ListColumn
    Dim sourceColumn As ListColumn
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim targetCell As Range
    Dim checkBox As Object
    Dim item As Object
    Dim isIncluded As Boolean

    For rowIndex = ws.CheckBoxes.Count To 1 Step -1
        Set item = ws.CheckBoxes(rowIndex)
        If Left$(CStr(item.Name), Len(PENDING_SOURCE_CHECK_PREFIX)) = PENDING_SOURCE_CHECK_PREFIX Then item.Delete
    Next rowIndex

    Set includeColumn = table.ListColumns("Include in Index?")
    Set sourceColumn = table.ListColumns("Source File")
    If table.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To table.ListRows.Count
        sourceFile = Trim$(CStr(sourceColumn.DataBodyRange.Cells(rowIndex, 1).Value))
        Set targetCell = includeColumn.DataBodyRange.Cells(rowIndex, 1)

        If Len(sourceFile) > 0 Then
            isIncluded = NMDC_SourceCheckedValue(targetCell.Value)
            targetCell.Value = isIncluded
            targetCell.NumberFormat = ";;;"
            targetCell.HorizontalAlignment = xlCenter

            Set checkBox = ws.CheckBoxes.Add( _
                targetCell.Left + (targetCell.Width - 13) / 2, _
                targetCell.Top + (targetCell.Height - 13) / 2, _
                13, 13)
            checkBox.Name = PENDING_SOURCE_CHECK_PREFIX & Format$(rowIndex, "0000")
            checkBox.Caption = ""
            checkBox.LinkedCell = "'" & ws.Name & "'!" & targetCell.Address
            checkBox.OnAction = ""
            checkBox.Value = IIf(isIncluded, xlOn, xlOff)
            checkBox.Placement = xlMoveAndSize
            checkBox.PrintObject = False
        Else
            targetCell.ClearContents
        End If
    Next rowIndex
    Exit Sub

Handler:
    NMDC_LogError "PENDING_SOURCE_CHECKBOX_BUILD_ERROR", _
        "Excel could not build the Pending Update source checkboxes.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_SaveSourceSelections()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim projectNo As String
    Dim ownerNote As String
    Dim includeValue As Boolean
    Dim csvText As String
    Dim decisionsPath As String
    Dim args As String
    Dim answer As VbMsgBoxResult
    Dim savedCount As Long

    Set ws = ThisWorkbook.Worksheets("Pending Update")
    Set table = ws.ListObjects(SOURCE_PANEL_TABLE)

    answer = MsgBox("Save these source choices and restage the proposal?" & vbCrLf & vbCrLf & _
                    "Checked = include source" & vbCrLf & _
                    "Unchecked = exclude source" & vbCrLf & vbCrLf & _
                    "No source workbook will be edited or deleted, and the approved index will not change until Approve Update is used.", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer <> vbYes Then Exit Sub

    csvText = NMDC_CheckboxCsvField("Source File") & "," & _
              NMDC_CheckboxCsvField("Include in Index?") & "," & _
              NMDC_CheckboxCsvField("Owner Note") & "," & _
              NMDC_CheckboxCsvField("Project No.") & vbLf

    If Not table.DataBodyRange Is Nothing Then
        For rowIndex = 1 To table.ListRows.Count
            sourceFile = Trim$(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value))
            If Len(sourceFile) > 0 Then
                includeValue = NMDC_SourceCheckedValue(table.ListColumns("Include in Index?").DataBodyRange.Cells(rowIndex, 1).Value)
                ownerNote = Trim$(CStr(table.ListColumns("Owner Note").DataBodyRange.Cells(rowIndex, 1).Value))
                projectNo = Trim$(CStr(table.ListColumns("Project No.").DataBodyRange.Cells(rowIndex, 1).Value))
                csvText = csvText & NMDC_CheckboxCsvField(sourceFile) & "," & _
                          NMDC_CheckboxCsvField(IIf(includeValue, "TRUE", "FALSE")) & "," & _
                          NMDC_CheckboxCsvField(ownerNote) & "," & _
                          NMDC_CheckboxCsvField(projectNo) & vbLf
                savedCount = savedCount + 1
            End If
        Next rowIndex
    End If

    If savedCount = 0 Then
        MsgBox "There are no source rows to save yet. Run a scan first.", vbInformation, "NMDC Document Index"
        Exit Sub
    End If

    decisionsPath = NMDC_ExchangePath() & "\source_selection_decisions.csv"
    NMDC_WriteCheckboxUtf8 decisionsPath, csvText
    args = "--selections-file " & NMDC_Quote(decisionsPath)

    If NMDC_RunEngine("save-source-selections", args) <> 0 Then
        MsgBox "The source choices could not be saved. Please review Error Log.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    MsgBox CStr(savedCount) & " source choice(s) saved. The proposal will now be restaged.", _
           vbInformation, "NMDC Document Index"
    NMDC_UpdateChangedFilesFast
    Exit Sub

Handler:
    NMDC_LogError "PENDING_SOURCE_SAVE_ERROR", _
        "Excel could not save the Pending Update source choices.", _
        Err.Number & " - " & Err.Description
    MsgBox "Source choices could not be saved. Please review Error Log.", _
           vbExclamation, "NMDC Document Index"
End Sub

Public Sub NMDC_CheckAllSources()
    NMDC_SetAllPendingSourceCheckboxes True
End Sub

Public Sub NMDC_UncheckAllSources()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Uncheck every source workbook?" & vbCrLf & vbCrLf & _
                    "Nothing changes until you click Save Source Choices & Restage.", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer = vbYes Then NMDC_SetAllPendingSourceCheckboxes False
End Sub

Private Sub NMDC_SetAllPendingSourceCheckboxes(ByVal checkedValue As Boolean)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim targetCell As Range
    Dim checkBox As Object

    Set ws = ThisWorkbook.Worksheets("Pending Update")
    Set table = ws.ListObjects(SOURCE_PANEL_TABLE)
    If table.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To table.ListRows.Count
        sourceFile = Trim$(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value))
        If Len(sourceFile) > 0 Then
            Set targetCell = table.ListColumns("Include in Index?").DataBodyRange.Cells(rowIndex, 1)
            targetCell.Value = checkedValue
            targetCell.NumberFormat = ";;;"
            Set checkBox = Nothing
            On Error Resume Next
            Set checkBox = ws.CheckBoxes(PENDING_SOURCE_CHECK_PREFIX & Format$(rowIndex, "0000"))
            On Error GoTo Handler
            If Not checkBox Is Nothing Then checkBox.Value = IIf(checkedValue, xlOn, xlOff)
        End If
    Next rowIndex
    Exit Sub

Handler:
    NMDC_LogError "PENDING_SOURCE_ALL_ERROR", _
        "Excel could not update all source checkboxes.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_SourceCheckboxClicked()
    ' Compatibility stub for any stale workbook control created by an older build.
    ' Fresh workbooks no longer assign this macro to individual checkboxes.
    On Error Resume Next
    NMDC_GoToSheet "Pending Update"
    NMDC_RebuildPendingSourcePanel
    MsgBox "Source selection is now handled directly on Pending Update. Use the checkbox panel on the right, then Save Source Choices & Restage.", _
           vbInformation, "NMDC Document Index"
    On Error GoTo 0
End Sub

Private Sub NMDC_RetireSeparateSourceSelectionUI()
    On Error Resume Next

    Dim legacyWs As Worksheet
    Dim homeWs As Worksheet
    Dim customButton As Shape
    Dim area As Range

    Set legacyWs = ThisWorkbook.Worksheets("Source Selection")
    If Not legacyWs Is Nothing Then legacyWs.Visible = xlSheetVeryHidden

    Set homeWs = ThisWorkbook.Worksheets("Home")
    homeWs.Shapes("NMDC_Action_19").Delete
    homeWs.Range("A40:D41").ClearContents

    Set customButton = Nothing
    Set customButton = homeWs.Shapes("NMDC_Action_20")
    If Not customButton Is Nothing Then
        Set area = homeWs.Range("A40:D41")
        customButton.Left = area.Left + 2
        customButton.Top = area.Top + 2
        customButton.Width = area.Width - 4
        customButton.Height = area.Height - 4
    End If

    On Error GoTo 0
End Sub

Public Sub NMDC_EnsureOwnerEnhancements()
    If mOwnerEnhancementsReady Then
        On Error Resume Next
        Application.Run "NMDC_LiveFilterWake"
        On Error GoTo 0
        Exit Sub
    End If

    On Error Resume Next
    Application.Run "NMDC_LiveFilterWake"
    Application.Run "NMDC_CustomFieldsInitialize"
    On Error GoTo 0
    mOwnerEnhancementsReady = True
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
