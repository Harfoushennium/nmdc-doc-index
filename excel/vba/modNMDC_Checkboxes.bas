Attribute VB_Name = "modNMDC_Checkboxes"
Option Explicit

Private Const SOURCE_PANEL_TABLE As String = "SourceSelection"
Private mOwnerEnhancementsReady As Boolean

' Source selection is integrated into Pending Update.
' Binary include/exclude choices use the modern Microsoft 365 in-cell Checkbox
' control (Range.CellControl.SetCheckbox).  No Form Control / ActiveX checkbox
' and no per-checkbox macro is required.

Public Sub NMDC_RebuildSourceSelectionCheckboxes()
    NMDC_RebuildPendingSourcePanel
End Sub

' Move the owner panel out of the incoming PendingUpdate range before Excel
' resizes that table. Snapshot its values first; the caller has already
' journaled owner edits against the current proposal.
Public Function NMDC_PreparePendingSourcePanelForResize(ByVal pendingTable As ListObject, _
                                                        ByVal incomingDataRows As Long) As Boolean
    On Error GoTo Handler
    Dim requiredRows As Long
    requiredRows = pendingTable.ListRows.Count
    If incomingDataRows > requiredRows Then requiredRows = incomingDataRows
    If requiredRows < 1 Then requiredRows = 1
    NMDC_PreparePendingSourcePanelForResize = NMDC_MovePendingSourcePanelToHeaderRow( _
        pendingTable.Parent, pendingTable.HeaderRowRange.Row + requiredRows + 4)
    Exit Function
Handler:
    NMDC_LogError "PENDING_SOURCE_PANEL_MOVE_ERROR", _
        "Excel could not determine a safe source-panel position before refreshing Pending Update.", _
        Err.Number & " - " & Err.Description
    NMDC_PreparePendingSourcePanelForResize = False
End Function

Private Function NMDC_MovePendingSourcePanelToHeaderRow(ByVal ws As Worksheet, ByVal targetHeaderRow As Long) As Boolean
    On Error GoTo Handler

    Dim panel As ListObject
    Dim oldRange As Range
    Dim oldHeaders As Variant
    Dim oldData As Variant
    Dim oldHeaderRow As Long
    Dim oldRowCount As Long
    Dim oldEvents As Boolean
    Dim errNumber As Long
    Dim errDescription As String
    Dim snapshotReady As Boolean
    Dim failedRange As Range

    oldEvents = Application.EnableEvents
    Set panel = Nothing
    On Error Resume Next
    Set panel = ws.ListObjects(SOURCE_PANEL_TABLE)
    On Error GoTo Handler
    If panel Is Nothing Then
        NMDC_MovePendingSourcePanelToHeaderRow = True
        Exit Function
    End If

    If panel.HeaderRowRange.Row = targetHeaderRow And panel.Range.Column = 1 Then
        NMDC_MovePendingSourcePanelToHeaderRow = True
        Exit Function
    End If

    Application.EnableEvents = False
    oldHeaderRow = panel.HeaderRowRange.Row
    Set oldRange = panel.Range
    oldHeaders = panel.HeaderRowRange.Value2
    oldRowCount = panel.ListRows.Count
    If Not panel.DataBodyRange Is Nothing Then oldData = panel.DataBodyRange.Value2
    snapshotReady = True

    NMDC_ClearPendingPanelDecoration ws, oldHeaderRow
    panel.Unlist
    oldRange.ClearContents RemoveControls:=True
    oldRange.ClearFormats

    Set panel = NMDC_CreatePendingSourceTable(ws, targetHeaderRow, oldHeaders, oldData, oldRowCount)
    NMDC_FormatPendingSourcePanel ws, panel
    NMDC_ApplyNativeSourceCheckboxes panel
    Application.EnableEvents = oldEvents
    NMDC_MovePendingSourcePanelToHeaderRow = True
    Exit Function

Handler:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    Set panel = Nothing
    Set panel = ws.ListObjects(SOURCE_PANEL_TABLE)
    If Not panel Is Nothing Then
        Set failedRange = panel.Range
        panel.Unlist
        failedRange.ClearContents RemoveControls:=True
        failedRange.ClearFormats
    End If
    If snapshotReady Then
        NMDC_ClearPendingPanelDecoration ws, targetHeaderRow
        Set panel = NMDC_CreatePendingSourceTable(ws, oldHeaderRow, oldHeaders, oldData, oldRowCount)
        NMDC_FormatPendingSourcePanel ws, panel
        NMDC_ApplyNativeSourceCheckboxes panel
    End If
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    NMDC_LogError "PENDING_SOURCE_PANEL_MOVE_ERROR", _
        "Excel could not move the source panel clear of the refreshed Pending Update rows.", _
        errNumber & " - " & errDescription
    NMDC_MovePendingSourcePanelToHeaderRow = False
End Function

Private Function NMDC_CreatePendingSourceTable(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                               ByVal headers As Variant, ByVal data As Variant, _
                                               ByVal dataRows As Long) As ListObject
    Dim bodyRows As Long
    Dim columnCount As Long
    Dim sourceRange As Range
    Dim table As ListObject
    bodyRows = dataRows
    If bodyRows < 1 Then bodyRows = 1
    columnCount = UBound(headers, 2)
    Set sourceRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow + bodyRows, columnCount))
    sourceRange.Rows(1).Value2 = headers
    Set table = ws.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
    table.Name = SOURCE_PANEL_TABLE
    table.TableStyle = "TableStyleMedium2"
    If dataRows > 0 Then
        table.DataBodyRange.Value2 = data
    ElseIf table.ListRows.Count = 0 Then
        table.ListRows.Add
    End If
    Set NMDC_CreatePendingSourceTable = table
End Function

Private Sub NMDC_ClearPendingPanelDecoration(ByVal ws As Worksheet, ByVal headerRow As Long)
    On Error Resume Next
    ws.Range(ws.Cells(headerRow - 2, 1), ws.Cells(headerRow - 2, 8)).UnMerge
    ws.Range(ws.Cells(headerRow - 2, 1), ws.Cells(headerRow - 2, 8)).ClearContents
    ws.Range(ws.Cells(headerRow - 2, 1), ws.Cells(headerRow - 2, 8)).ClearFormats
    ws.Range(ws.Cells(headerRow - 1, 1), ws.Cells(headerRow - 1, 8)).ClearContents
    ws.Range(ws.Cells(headerRow - 1, 1), ws.Cells(headerRow - 1, 8)).ClearFormats
    ws.Shapes("NMDC_Pending_Exclude").Delete
    ws.Shapes("NMDC_Pending_SourceList").Delete
    ws.Shapes("NMDC_Pending_SourceSave").Delete
    ws.Shapes("NMDC_Pending_SourceAll").Delete
    ws.Shapes("NMDC_Pending_SourceNone").Delete
    ws.Shapes("NMDC_Pending_ReviewDrafts").Delete
    On Error GoTo 0
End Sub

Public Sub NMDC_RebuildPendingSourcePanel()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim oldEvents As Boolean

    oldEvents = Application.EnableEvents
    If Not NMDC_ReviewDraftStoreHealthy() Or Not NMDC_ReviewDraftCaptureCurrentTables() Then Exit Sub
    Application.EnableEvents = False
    NMDC_EnsureOwnerEnhancements
    NMDC_RetireSeparateSourceSelectionUI

    Set ws = ThisWorkbook.Worksheets("Pending Update")
    Set table = NMDC_EnsurePendingSourceTable(ws)
    If table Is Nothing Then GoTo CleanExit

    If NMDC_FileExists(NMDC_ExchangePath() & "\source_selection.csv") Then
        If Not NMDC_LoadSourceSelectionPanel(table, NMDC_ExchangePath() & "\source_selection.csv") Then GoTo CleanExit
    End If

    NMDC_FormatPendingSourcePanel ws, table
    NMDC_ApplyNativeSourceCheckboxes table
    Call NMDC_ReviewDraftsRestoreCurrent()
CleanExit:
    Application.EnableEvents = oldEvents
    Exit Sub

Handler:
    On Error Resume Next
    Application.EnableEvents = oldEvents
    On Error GoTo 0
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
    Dim pendingTable As ListObject
    Dim headerRow As Long
    Dim firstColumn As Long

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

        Set pendingTable = Nothing
        On Error Resume Next
        Set pendingTable = ws.ListObjects("PendingUpdate")
        On Error GoTo Handler
        headerRow = 8
        If Not pendingTable Is Nothing Then headerRow = pendingTable.Range.Row + pendingTable.Range.Rows.Count + 3
        firstColumn = 1

        For i = LBound(headers) To UBound(headers)
            ws.Cells(headerRow, firstColumn + i - LBound(headers)).Value = headers(i)
        Next i

        Set sourceRange = ws.Range(ws.Cells(headerRow, firstColumn), ws.Cells(headerRow + 1, firstColumn + 7))
        Set table = ws.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
        table.Name = SOURCE_PANEL_TABLE
        table.TableStyle = "TableStyleMedium2"
        If table.ListRows.Count = 0 Then table.ListRows.Add
    Else
        Set pendingTable = Nothing
        On Error Resume Next
        Set pendingTable = ws.ListObjects("PendingUpdate")
        On Error GoTo Handler
        If Not pendingTable Is Nothing Then
            headerRow = pendingTable.Range.Row + pendingTable.Range.Rows.Count + 3
            If table.Range.Row <> headerRow Or table.Range.Column <> 1 Then
                If Not NMDC_MovePendingSourcePanelToHeaderRow(ws, headerRow) Then _
                    Err.Raise vbObjectError + 240, , "The source-selection panel could not be moved below Pending Update."
                Set table = ws.ListObjects(SOURCE_PANEL_TABLE)
            End If
        End If
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

    If Not NMDC_ParseCsvFile(csvPath, headers, data, dataCount, columnCount, True) Then Exit Function

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
        If table.DataBodyRange Is Nothing Then table.ListRows.Add
        table.DataBodyRange.Value2 = data
    Else
        If table.ListRows.Count = 0 Then table.ListRows.Add
        If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.ClearContents
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
    Dim headerRow As Long
    Dim titleRow As Long
    Dim buttonRow As Long

    headerRow = table.HeaderRowRange.Row
    titleRow = headerRow - 2
    buttonRow = headerRow - 1

    ' The source panel is compact and directly below PendingUpdate, not a
    ' far-right block that makes the sheet excessively wide.
    On Error Resume Next
    ws.Range(ws.Cells(titleRow, 1), ws.Cells(titleRow, 8)).UnMerge
    ws.Range(ws.Cells(titleRow, 1), ws.Cells(titleRow, 8)).ClearContents
    ws.Range(ws.Cells(titleRow, 1), ws.Cells(titleRow, 8)).Merge
    On Error GoTo Handler
    With ws.Range(ws.Cells(titleRow, 1), ws.Cells(titleRow, 8))
        .Cells(1, 1).Value = "SOURCE SELECTION - native Excel checkboxes: checked = INCLUDE, unchecked = EXCLUDE. Project No. identifies each source. Save once after all choices."
        .Interior.Color = RGB(232, 241, 247)
        .Font.Name = "Aptos"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(31, 70, 90)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .RowHeight = 42
    End With

    On Error Resume Next
    ws.Shapes("NMDC_Pending_Exclude").Delete
    ws.Shapes("NMDC_Pending_SourceList").Delete
    ws.Shapes("NMDC_Pending_SourceSave").Delete
    ws.Shapes("NMDC_Pending_SourceAll").Delete
    ws.Shapes("NMDC_Pending_SourceNone").Delete
    On Error GoTo Handler

    Set area = ws.Range(ws.Cells(buttonRow, 1), ws.Cells(buttonRow, 3))
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_SourceSave"
    button.OnAction = "NMDC_SaveSourceSelections"
    button.TextFrame.Characters.Text = "Save Source Choices" & vbLf & "& Restage"
    NMDC_FormatSourceButton button, RGB(20, 108, 148)

    Set area = ws.Range(ws.Cells(buttonRow, 4), ws.Cells(buttonRow, 5))
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_SourceAll"
    button.OnAction = "NMDC_CheckAllSources"
    button.TextFrame.Characters.Text = "Check All"
    NMDC_FormatSourceButton button, RGB(46, 125, 50)

    Set area = ws.Range(ws.Cells(buttonRow, 6), ws.Cells(buttonRow, 8))
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_SourceNone"
    button.OnAction = "NMDC_UncheckAllSources"
    button.TextFrame.Characters.Text = "Uncheck All"
    NMDC_FormatSourceButton button, RGB(198, 40, 40)

    table.ListColumns("Include in Index?").Range.ColumnWidth = 13
    table.ListColumns("Project No.").Range.ColumnWidth = 16
    table.ListColumns("Source File").Range.ColumnWidth = 42
    table.ListColumns("Owner Note").Range.ColumnWidth = 28

    On Error Resume Next
    ws.Range("A4:I4").UnMerge
    ws.Range("A4:I4").ClearContents
    ws.Range("A4:I4").Merge
    ws.Range("A4").Value = _
        "REVIEW WORKFLOW. Review staged changes in the first table. Use the native source checkboxes in the source-selection section below it, then Save Source Choices & Restage. Approve/Hold/Reject applies to the remaining whole proposal."
    ThisWorkbook.Worksheets("Home").Range("E34").Value = _
        "Pending Update contains the source-selection section below the staged changes table. Checked = include; unchecked = exclude. Save Source Choices & Restage once when finished."
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

    Set includeColumn = table.ListColumns("Include in Index?")
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
           "Please close Excel, reopen the workbook, and try Pending Update again." & vbCrLf & vbCrLf & _
           "Technical detail: " & CStr(errNumber) & " - " & errDescription, _
           vbExclamation, "NMDC Document Index"
    On Error GoTo 0
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

    Set ws = ThisWorkbook.Worksheets("Pending Update")
    Set table = ws.ListObjects(SOURCE_PANEL_TABLE)

    answer = MsgBox("Save these source choices and restage the proposal?" & vbCrLf & vbCrLf & _
                    "Checked = include source" & vbCrLf & _
                    "Unchecked = exclude source" & vbCrLf & vbCrLf & _
                    "Source workbooks are never edited or deleted. The approved index remains unchanged until Approve Update.", _
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

    ' Persist the approval guard before changing owner configuration. If this
    ' write fails, no config change can leave the displayed proposal approvable.
    NMDC_ReviewDraftMarkSourceRestage NMDC_DisplayedPendingRunId()

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
    NMDC_SetAllPendingSourceChoices True
End Sub

Public Sub NMDC_UncheckAllSources()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Uncheck every source workbook?" & vbCrLf & vbCrLf & _
                    "Nothing changes until you click Save Source Choices & Restage.", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer = vbYes Then NMDC_SetAllPendingSourceChoices False
End Sub

Private Sub NMDC_SetAllPendingSourceChoices(ByVal checkedValue As Boolean)
    On Error GoTo Handler

    Dim table As ListObject
    Dim rowIndex As Long
    Dim sourceFile As String
    Dim oldEvents As Boolean

    Set table = ThisWorkbook.Worksheets("Pending Update").ListObjects(SOURCE_PANEL_TABLE)
    If table.DataBodyRange Is Nothing Then Exit Sub

    oldEvents = Application.EnableEvents
    Application.EnableEvents = False
    For rowIndex = 1 To table.ListRows.Count
        sourceFile = Trim$(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value))
        If Len(sourceFile) > 0 Then
            table.ListColumns("Include in Index?").DataBodyRange.Cells(rowIndex, 1).Value = checkedValue
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
    NMDC_GoToSheet "Pending Update"
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
