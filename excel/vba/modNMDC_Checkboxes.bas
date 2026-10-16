Attribute VB_Name = "modNMDC_Checkboxes"
Option Explicit

Private Const SOURCE_PANEL_TABLE As String = "FileReview"
Private mOwnerEnhancementsReady As Boolean

' Proposal acceptance is on File Review; persistent source scope is an engine configuration concern.
' Binary accept/defer choices use the modern Microsoft 365 in-cell Checkbox
' control (Range.CellControl.SetCheckbox).  No Form Control / ActiveX checkbox
' and no per-checkbox macro is required.

Public Sub NMDC_RebuildSourceSelectionCheckboxes()
    NMDC_RebuildPendingSourcePanel
End Sub

Public Function NMDC_PreparePendingSourcePanelForResize(ByVal pendingTable As ListObject, ByVal incomingDataRows As Long) As Boolean
    ' File Review has its own worksheet; Change Details can resize independently.
    NMDC_PreparePendingSourcePanelForResize = True
End Function

Public Function NMDC_RebuildPendingSourcePanel() As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim oldEvents As Boolean

    oldEvents = Application.EnableEvents
    If Not NMDC_ReviewDraftStoreHealthy() Or Not NMDC_ReviewDraftCaptureCurrentTables() Then Exit Function
    Application.EnableEvents = False
    NMDC_EnsureOwnerEnhancements
    NMDC_RetireSeparateSourceSelectionUI

    Set ws = ThisWorkbook.Worksheets("File Review")
    Set table = NMDC_EnsurePendingSourceTable(ws)
    If table Is Nothing Then GoTo CleanExit

    If Not NMDC_LoadSourceSelectionPanel(table, NMDC_ExchangePath() & "\file_review.csv") Then GoTo CleanExit

    NMDC_FormatPendingSourcePanel ws, table
    NMDC_ApplyNativeSourceCheckboxes table
    If Not NMDC_ReviewDraftsRestoreCurrent() Then GoTo CleanExit
    NMDC_RebuildPendingSourcePanel = True
CleanExit:
    Application.EnableEvents = oldEvents
    Exit Function

Handler:
    On Error Resume Next
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    NMDC_LogError "PENDING_SOURCE_PANEL_ERROR", _
        "Excel could not prepare source selection inside File Review.", _
        Err.Number & " - " & Err.Description
End Function

Private Function NMDC_EnsurePendingSourceTable(ByVal ws As Worksheet) As ListObject
    Set NMDC_EnsurePendingSourceTable = ws.ListObjects(SOURCE_PANEL_TABLE)
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

    If Not NMDC_ParseCsvFile(csvPath, headers, data, dataCount, columnCount, True) Then Exit Function
    NMDC_ResetTableViewForRefresh table
    NMDC_ClearGeneratedTableHyperlinks table

    tableRows = dataCount
    If tableRows < 1 Then tableRows = 1
    headerRow = table.HeaderRowRange.Row
    firstColumn = table.Range.Column

    If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.ClearContents RemoveControls:=True

    Set targetRange = table.Parent.Range( _
        table.Parent.Cells(headerRow, firstColumn), _
        table.Parent.Cells(headerRow + tableRows, firstColumn + columnCount - 1))
    table.Resize targetRange
    table.HeaderRowRange.Value2 = headers

    If dataCount > 0 Then
        If table.DataBodyRange Is Nothing Then table.ListRows.Add
        ' Set text before assignment: formatting afterward cannot undo Excel's
        ' numeric coercion of project identifiers on the first proposal load.
        table.DataBodyRange.NumberFormat = "@"
        table.DataBodyRange.Value2 = data
    Else
        If table.ListRows.Count = 0 Then table.ListRows.Add
        If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.ClearContents
    End If

    NMDC_ApplyTypedFormatting table
    NMDC_ActivateSourceLinks table
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
    Dim labels As Variant, actions As Variant, i As Long
    Dim area As Range, button As Shape
    Const BUTTON_WIDTH As Double = 128
    Const BUTTON_GAP As Double = 6
    labels = Array("Save File Choices", "Check All", "Uncheck All", "Approve Update", "View Change Details", "Review Saved Drafts")
    actions = Array("NMDC_SaveSourceSelections", "NMDC_CheckAllSources", "NMDC_UncheckAllSources", "NMDC_ApproveUpdate", "NMDC_ViewChangeDetails", "NMDC_ReviewDraftOpenReport")
    ws.Range("A1").Value = "File Review"
    ' This owned instruction row stays above the Table. Clear before merging;
    ' no table or owner-input cells are in this range.
    ws.Range("A4:H4").UnMerge
    ws.Range("A4:H4").ClearContents
    ws.Range("A4:H4").Merge
    ws.Range("A4").Value = "Checked: accept this file's update. Unchecked: keep its approved data for this proposal. Save File Choices before approval."
    With ws.Range("A4:H4")
        .WrapText = True
        .Font.Name = "Aptos"
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 60
    End With
    table.ListColumns("Accept Update?").Range.ColumnWidth = 15
    table.ListColumns("Project No.").Range.ColumnWidth = 16
    table.ListColumns("Source File").Range.ColumnWidth = 48
    table.ListColumns("Change Type").Range.ColumnWidth = 12
    table.ListColumns("Added").Range.ColumnWidth = 8
    table.ListColumns("Modified").Range.ColumnWidth = 8
    table.ListColumns("Removed").Range.ColumnWidth = 8
    table.ListColumns("Owner Note").Range.ColumnWidth = 32
    ws.Rows(2).RowHeight = 30
    Set area = ws.Range("A2")
    For i = 0 To UBound(labels)
        On Error Resume Next
        ws.Shapes("NMDC_FileReview_" & CStr(i)).Delete
        On Error GoTo 0
        ' Keep all actions in one compact toolbar, independent of table widths.
        Set button = ws.Shapes.AddShape(5, area.Left + i * (BUTTON_WIDTH + BUTTON_GAP), area.Top, BUTTON_WIDTH, area.Height)
        button.Name = "NMDC_FileReview_" & CStr(i)
        button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & actions(i)
        button.TextFrame.Characters.Text = labels(i)
        NMDC_FormatSourceButton button, RGB(20, 108, 148)
    Next i
End Sub

Public Sub NMDC_ViewChangeDetails()
    NMDC_GoToSheet "Change Details"
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
        .Placement = xlFreeFloating
    End With
End Sub

Private Sub NMDC_ApplyNativeSourceCheckboxes(ByVal table As ListObject)
    On Error GoTo CheckboxRequiredFailed

    Dim includeColumn As ListColumn
    Dim sourceColumn As ListColumn
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim targetCell As Range
    Dim anySource As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    Set includeColumn = table.ListColumns("Accept Update?")
    Set sourceColumn = table.ListColumns("Source File")
    If table.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To table.ListRows.Count
        sourceFile = Trim$(CStr(sourceColumn.DataBodyRange.Cells(rowIndex, 1).Value))
        Set targetCell = includeColumn.DataBodyRange.Cells(rowIndex, 1)
        If Len(sourceFile) > 0 Then
            targetCell.Value = NMDC_SourceCheckedValue(targetCell.Value)
            anySource = True
        Else
            targetCell.ClearContents
        End If
    Next rowIndex

    If anySource Then
        ' Owner environment is Microsoft 365 with native in-cell checkboxes.
        ' Native checkbox rendering is therefore a production requirement,
        ' not an optional TRUE/FALSE fallback.
        includeColumn.DataBodyRange.CellControl.SetCheckbox
        includeColumn.DataBodyRange.HorizontalAlignment = xlCenter
    End If
    Exit Sub

CheckboxRequiredFailed:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    NMDC_LogError "NATIVE_CHECKBOX_REQUIRED", _
        "Excel could not create the required Microsoft 365 source-selection checkboxes.", _
        errNumber & " - " & errDescription
    MsgBox "The required Microsoft 365 source-selection checkboxes could not be created." & vbCrLf & vbCrLf & _
           "This production build does not fall back to TRUE/FALSE text." & vbCrLf & _
           "Please close Excel, reopen the workbook, and try File Review again." & vbCrLf & vbCrLf & _
           "Technical detail: " & CStr(errNumber) & " - " & errDescription, _
           vbExclamation, "NMDC Document Index"
    On Error GoTo 0
    Err.Raise vbObjectError + 243, , "Required native File Review checkbox creation failed. " & errDescription
End Sub

Public Sub NMDC_SaveSourceSelections()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim projectNo As String
    If NMDC_AsyncIsBusy() Then
        MsgBox "Wait for the background update to finish before saving source choices.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If
    If Not NMDC_ReviewDraftStoreHealthy() Or Not NMDC_ReviewDraftCaptureCurrentTables() Then Exit Sub
    Dim ownerNote As String
    Dim includeValue As Boolean
    Dim csvText As String
    Dim decisionsPath As String
    Dim args As String
    Dim answer As VbMsgBoxResult
    Dim savedCount As Long

    Set ws = ThisWorkbook.Worksheets("File Review")
    Set table = ws.ListObjects(SOURCE_PANEL_TABLE)

    answer = MsgBox("Save file choices for this proposal?" & vbCrLf & vbCrLf & _
                    "Checked = accept this file update" & vbCrLf & _
                    "Unchecked = keep previously approved data" & vbCrLf & vbCrLf & _
                    "Source workbooks are never edited or deleted. The approved index remains unchanged until Approve Update.", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer <> vbYes Then Exit Sub

    csvText = NMDC_CheckboxCsvField("Source File") & "," & _
              NMDC_CheckboxCsvField("Accept Update?") & "," & _
              NMDC_CheckboxCsvField("Owner Note") & "," & _
              NMDC_CheckboxCsvField("Project No.") & vbLf

    If Not table.DataBodyRange Is Nothing Then
        For rowIndex = 1 To table.ListRows.Count
            sourceFile = Trim$(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value))
            If Len(sourceFile) > 0 Then
                includeValue = NMDC_SourceCheckedValue(table.ListColumns("Accept Update?").DataBodyRange.Cells(rowIndex, 1).Value)
                ownerNote = CStr(table.ListColumns("Owner Note").DataBodyRange.Cells(rowIndex, 1).Value)
                projectNo = Trim$(CStr(table.ListColumns("Project No.").DataBodyRange.Cells(rowIndex, 1).Value))
                csvText = csvText & NMDC_CheckboxCsvField(sourceFile) & "," & _
                          NMDC_CheckboxCsvField(IIf(includeValue, "TRUE", "FALSE")) & "," & _
                          NMDC_CheckboxCsvField(ownerNote) & "," & _
                          NMDC_CheckboxCsvField(projectNo) & vbLf
                savedCount = savedCount + 1
            End If
        Next rowIndex
    End If

    If Len(NMDC_DisplayedPendingRunId()) = 0 Then Exit Sub

    decisionsPath = NMDC_ExchangePath() & "\file_review_decisions.csv"
    NMDC_WriteCheckboxUtf8 decisionsPath, csvText
    args = "--run-id " & NMDC_Quote(NMDC_DisplayedPendingRunId()) & " --choices-file " & NMDC_Quote(decisionsPath)

    If NMDC_RunEngine("save-file-review", args) <> 0 Then
        MsgBox "File choices could not be saved. Review Error Log.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If
    If NMDC_RunEngine("export-excel") <> 0 Then Exit Sub
    ' Refresh compares the live cells to the saved exchange baseline. Matching
    ' drafts retire; edits made while the engine ran remain draft and block approval.
    If Not NMDC_RefreshReviewDataFast() Then Exit Sub
    MsgBox CStr(savedCount) & " file choice(s) saved for this proposal. Nothing is approved yet.", vbInformation, "NMDC Document Index"
    Exit Sub

Handler:
    NMDC_LogError "PENDING_SOURCE_SAVE_ERROR", _
        "Excel could not save the File Review source choices.", _
        Err.Number & " - " & Err.Description
    MsgBox "Source choices could not be saved. Please review Error Log.", _
           vbExclamation, "NMDC Document Index"
End Sub

Public Sub NMDC_CheckAllSources()
    NMDC_SetAllPendingSourceChoices True
End Sub

Public Sub NMDC_UncheckAllSources()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Uncheck every source workbook?" & vbCrLf & vbCrLf & _
                    "Nothing changes until you click Save File Choices.", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer = vbYes Then NMDC_SetAllPendingSourceChoices False
End Sub

Private Sub NMDC_SetAllPendingSourceChoices(ByVal checkedValue As Boolean)
    On Error GoTo Handler

    Dim table As ListObject
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim oldEvents As Boolean

    Set table = ThisWorkbook.Worksheets("File Review").ListObjects(SOURCE_PANEL_TABLE)
    If table.DataBodyRange Is Nothing Then Exit Sub

    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    For rowIndex = 1 To table.ListRows.Count
        sourceFile = Trim$(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value))
        If Len(sourceFile) > 0 Then
            table.ListColumns("Accept Update?").DataBodyRange.Cells(rowIndex, 1).Value = checkedValue
        End If
    Next rowIndex
    Application.EnableEvents = oldEvents
    NMDC_ReviewDraftCaptureAllSourceChoices
    Exit Sub

Handler:
    On Error Resume Next
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    NMDC_LogError "PENDING_SOURCE_ALL_ERROR", _
        "Excel could not update all source checkboxes.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_SourceCheckboxClicked()
    ' Compatibility stub only. Fresh workbooks use native in-cell checkboxes,
    ' which directly store TRUE/FALSE and do not call a macro when clicked.
    On Error Resume Next
    NMDC_GoToSheet "File Review"
    NMDC_RebuildPendingSourcePanel
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
        Case "TRUE", "YES", "1", "ON", "CHECKED", "INCLUDE", "INCLUDED", "SELECTED"
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
