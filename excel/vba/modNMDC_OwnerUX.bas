Attribute VB_Name = "modNMDC_OwnerUX"
Option Explicit

Public Sub NMDC_ApplyOwnerUX()
    On Error GoTo Handler
    NMDC_ApplyHomeButtonGuides
    NMDC_ApplyHomeWorkflowGuide
    NMDC_ApplyPendingUpdateDecisionGuide
    NMDC_ApplySourceSelectionGuide
    NMDC_ApplyPlainRulesExperience
    Exit Sub
Handler:
    NMDC_LogError "OWNER_UX_ERROR", _
        "Excel could not apply all owner guidance controls.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_OpenSourceSelection()
    On Error GoTo Handler
    If NMDC_RunEngine("export-excel") <> 0 Then
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    If Not NMDC_RefreshSourceSelectionTable() Then
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    NMDC_GoToSheet "Source Selection"
    Exit Sub
Handler:
    NMDC_LogError "SOURCE_SELECTION_OPEN_ERROR", _
        "Excel could not open Source Selection.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_OpenRulesMappingsOwner()
    NMDC_GoToSheet "Rules & Mappings"
    NMDC_ApplyPlainRulesExperience
End Sub

Public Sub NMDC_ExcludeSelectedSource()
    NMDC_SetSelectedSourceChoice False
End Sub

Public Sub NMDC_IncludeSelectedSource()
    NMDC_SetSelectedSourceChoice True
End Sub

Private Sub NMDC_SetSelectedSourceChoice(ByVal includeSource As Boolean)
    On Error GoTo Handler

    Dim sourceFile As String
    Dim actionText As String
    Dim reason As String
    Dim args As String
    Dim answer As VbMsgBoxResult

    sourceFile = NMDC_SelectedSourceFile()
    If Len(sourceFile) = 0 Then
        MsgBox "Select a data row first in Pending Update or Source Selection." & vbCrLf & _
               "The source file is taken from the selected row.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    If includeSource Then
        actionText = "INCLUDE"
        answer = MsgBox("Include this source workbook in the index?" & vbCrLf & vbCrLf & _
                        sourceFile & vbCrLf & vbCrLf & _
                        "The proposal will be restaged. Nothing is approved automatically.", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index")
        If answer <> vbYes Then Exit Sub
    Else
        actionText = "EXCLUDE"
        answer = MsgBox("Exclude this entire source workbook from the index?" & vbCrLf & vbCrLf & _
                        sourceFile & vbCrLf & vbCrLf & _
                        "All records originating only from this source will be omitted from the next staged proposal. " & _
                        "The source workbook itself will NOT be changed or deleted.", _
                        vbExclamation + vbYesNo + vbDefaultButton2, "NMDC Document Index")
        If answer <> vbYes Then Exit Sub
        reason = InputBox("Optional reason for excluding this source:" & vbCrLf & _
                          "Example: Duplicate register / not authoritative / outside index scope", _
                          "Exclude Source", "Owner excluded during Pending Update review")
    End If

    args = "--source-file " & NMDC_Quote(sourceFile) & _
           " --action " & actionText & _
           " --reason " & NMDC_Quote(reason)

    If NMDC_RunEngine("set-source-selection", args) <> 0 Then
        MsgBox "The source selection could not be saved. Please review Error Log.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    MsgBox "Source choice saved: " & actionText & vbCrLf & vbCrLf & _
           sourceFile & vbCrLf & vbCrLf & _
           "The index will now restage the proposal using this source choice. Nothing is approved yet.", _
           vbInformation, "NMDC Document Index"
    NMDC_UpdateChangedFilesFast
    Exit Sub

Handler:
    NMDC_LogError "SOURCE_SELECTION_SAVE_ERROR", _
        "Excel could not save the selected source choice.", _
        Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_SelectedSourceFile() As String
    On Error GoTo Failed

    Dim selectedCell As Range
    Dim table As ListObject
    Dim hit As Range
    Dim rowIndex As Long

    If TypeName(Selection) <> "Range" Then Exit Function
    Set selectedCell = Selection.Cells(1, 1)

    For Each table In selectedCell.Worksheet.ListObjects
        If Not table.DataBodyRange Is Nothing Then
            Set hit = Intersect(selectedCell, table.DataBodyRange)
            If Not hit Is Nothing Then
                On Error Resume Next
                rowIndex = selectedCell.Row - table.DataBodyRange.Row + 1
                NMDC_SelectedSourceFile = Trim$(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value))
                On Error GoTo Failed
                If Len(NMDC_SelectedSourceFile) > 0 Then Exit Function
            End If
        End If
    Next table
    Exit Function
Failed:
    NMDC_SelectedSourceFile = ""
End Function

Public Function NMDC_RefreshSourceSelectionTable() As Boolean
    On Error GoTo Handler
    If Not NMDC_FileExists(NMDC_ExchangePath() & "\source_selection.csv") Then
        NMDC_RefreshSourceSelectionTable = True
        Exit Function
    End If
    NMDC_RefreshSourceSelectionTable = NMDC_LoadCsvToTable( _
        NMDC_ExchangePath() & "\source_selection.csv", "Source Selection", "SourceSelection")
    NMDC_ApplySourceSelectionGuide
    Exit Function
Handler:
    NMDC_LogError "SOURCE_SELECTION_REFRESH_ERROR", _
        "Excel could not refresh Source Selection.", Err.Number & " - " & Err.Description
    NMDC_RefreshSourceSelectionTable = False
End Function

Private Sub NMDC_ApplyHomeButtonGuides()
    Dim titles As Variant
    Dim guides As Variant
    Dim index As Long
    Dim shape As Shape

    titles = Array( _
        "Update Changed Files", "Full Rescan / Rebuild All", "Review Pending Update", _
        "Approve Update", "Hold Update", "Reject Update", _
        "Select Data Folder", "Review Flags", "Save Review Decisions", _
        "Configuration", "Rules & Mappings", "View Log", _
        "Flag Wrong Data", "Report Requirement / Problem", "Refresh Dashboard", _
        "Undo Last Approval", "Reset All Records", "Help")

    guides = Array( _
        "NORMAL USE - scan new/changed files only", _
        "HEAVY - rebuild every source from the local cache", _
        "Preview exactly what would change", _
        "Accept the whole staged proposal", _
        "Pause the proposal without changing approved data", _
        "Discard the staged proposal", _
        "Choose the folder that contains source registers", _
        "Review only genuine extraction/data exceptions", _
        "Save Review Flags choices and comments", _
        "View supported workbook settings", _
        "Edit plain-text classification rules", _
        "Open errors and diagnostic messages", _
        "Report the selected table value as incorrect", _
        "Record a problem or improvement request", _
        "Refresh status without running a new scan", _
        "Restore the previous approved index version", _
        "Clear indexed/staged records; sources remain untouched", _
        "Open the step-by-step workbook guide")

    For index = 0 To UBound(titles)
        Set shape = Nothing
        On Error Resume Next
        Set shape = ThisWorkbook.Worksheets("Home").Shapes("NMDC_Action_" & CStr(index + 1))
        On Error GoTo 0
        If Not shape Is Nothing Then
            shape.TextFrame.Characters.Text = titles(index) & vbLf & guides(index)
            shape.TextFrame.Characters.Font.Name = "Aptos"
            shape.TextFrame.Characters.Font.Size = 8.5
            shape.TextFrame.Characters.Font.Bold = True
            shape.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            shape.AlternativeText = titles(index) & ": " & guides(index)
        End If
    Next index

    ' Full Rescan is intentionally visually different because it is the heavy action.
    Set shape = Nothing
    On Error Resume Next
    Set shape = ThisWorkbook.Worksheets("Home").Shapes("NMDC_Action_2")
    On Error GoTo 0
    If Not shape Is Nothing Then
        shape.Fill.ForeColor.RGB = RGB(194, 139, 0)
        shape.Line.ForeColor.RGB = RGB(194, 139, 0)
    End If
End Sub

Private Sub NMDC_ApplyHomeWorkflowGuide()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Home")

    On Error Resume Next
    ws.Range("A30:L38").UnMerge
    On Error GoTo 0
    ws.Range("A30:L38").ClearContents

    ws.Range("A30:L30").Merge
    With ws.Range("A30:L30")
        .Value = "QUICK WORKFLOW - WHAT TO DO"
        .Interior.Color = RGB(20, 108, 148)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 22
    End With

    NMDC_SetGuideRow ws, 31, "1", "SELECT SOURCE", "Select Data Folder once. Keep the application/runtime outside OneDrive; the source folder may remain in OneDrive."
    NMDC_SetGuideRow ws, 32, "2", "SCAN", "For normal work press Update Changed Files. Full Rescan is only for deliberate rebuilds, parser/rule changes or troubleshooting."
    NMDC_SetGuideRow ws, 33, "3", "REVIEW", "Open Pending Update. It is a preview: no approved data changes until you press Approve Update."
    NMDC_SetGuideRow ws, 34, "4", "SOURCE DECISION", "If an entire workbook should not be indexed, open Source Selection or select its Pending Update row and choose Exclude Selected Source. Restage and review again."
    NMDC_SetGuideRow ws, 35, "5", "DECIDE", "Approve Update = accept the whole remaining proposal; Hold = postpone it; Reject = discard the staged proposal."
    NMDC_SetGuideRow ws, 36, "6", "EXCEPTIONS", "Review Flags is only for genuine extraction/data exceptions. For the current validated DATA, the expected count is zero."
    NMDC_SetGuideRow ws, 37, "TIP", "PERFORMANCE", "The scan runs in the background. A persistent local source cache avoids rereading unchanged OneDrive files on later scans."
    NMDC_SetGuideRow ws, 38, "TIP", "RULES", "Rules & Mappings uses normal text matching: CONTAINS, EXACT, STARTS WITH or ENDS WITH. No REGEX is needed for normal editing."
End Sub

Private Sub NMDC_SetGuideRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal stepText As String, ByVal titleText As String, ByVal detailText As String)
    ws.Range("A" & rowNo & ":B" & rowNo).Merge
    ws.Range("C" & rowNo & ":D" & rowNo).Merge
    ws.Range("E" & rowNo & ":L" & rowNo).Merge
    ws.Range("A" & rowNo).Value = stepText
    ws.Range("C" & rowNo).Value = titleText
    ws.Range("E" & rowNo).Value = detailText
    With ws.Range("A" & rowNo & ":L" & rowNo)
        .Interior.Color = RGB(247, 249, 252)
        .Font.Name = "Aptos"
        .Font.Size = 9
        .Font.Color = RGB(31, 41, 55)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 30
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(225, 230, 235)
        .Borders.Weight = xlHairline
    End With
    ws.Range("A" & rowNo).Font.Bold = True
    ws.Range("C" & rowNo).Font.Bold = True
End Sub

Private Sub NMDC_ApplyPendingUpdateDecisionGuide()
    Dim ws As Worksheet
    Dim area As Range
    Dim button As Shape

    Set ws = ThisWorkbook.Worksheets("Pending Update")
    On Error Resume Next
    ws.Range("A4:I4").UnMerge
    ws.Range("A4:I4").Merge
    On Error GoTo 0
    With ws.Range("A4:I4")
        .Value = "REVIEW WORKFLOW - READ ONLY TABLE. Review the proposed changes. To omit an entire source workbook, select any row from that source and click Exclude Selected Source. The proposal will restage. Approve/Hold/Reject always applies to the whole staged proposal, not one row."
        .Interior.Color = RGB(255, 247, 219)
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(96, 72, 0)
        .WrapText = True
        .RowHeight = 54
    End With

    On Error Resume Next
    ws.Shapes("NMDC_Pending_Exclude").Delete
    ws.Shapes("NMDC_Pending_SourceList").Delete
    On Error GoTo 0

    Set area = ws.Range("K2:M3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_Exclude"
    button.OnAction = "NMDC_ExcludeSelectedSource"
    button.TextFrame.Characters.Text = "Exclude Selected Source" & vbLf & "Omit this whole workbook"
    NMDC_FormatOwnerButton button, RGB(198, 40, 40)

    Set area = ws.Range("N2:P3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_SourceList"
    button.OnAction = "NMDC_OpenSourceSelection"
    button.TextFrame.Characters.Text = "Source Selection" & vbLf & "Include / restore sources"
    NMDC_FormatOwnerButton button, RGB(20, 108, 148)
End Sub

Private Sub NMDC_ApplySourceSelectionGuide()
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim area As Range
    Dim button As Shape

    Set ws = ThisWorkbook.Worksheets("Source Selection")
    On Error Resume Next
    ws.Range("A4:F4").UnMerge
    ws.Range("A4:F4").Merge
    On Error GoTo 0
    With ws.Range("A4:F4")
        .Value = "OWNER SOURCE SELECTION. INCLUDE = source participates in the index. EXCLUDE = source is intentionally omitted without deleting or editing the source workbook. Select a row, then use the Include/Exclude buttons. The proposal is restaged before approval."
        .Interior.Color = RGB(232, 241, 247)
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(31, 70, 90)
        .WrapText = True
        .RowHeight = 54
    End With

    On Error Resume Next
    ws.Shapes("NMDC_Source_Include").Delete
    ws.Shapes("NMDC_Source_Exclude").Delete
    On Error GoTo 0

    Set area = ws.Range("H2:J3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Source_Include"
    button.OnAction = "NMDC_IncludeSelectedSource"
    button.TextFrame.Characters.Text = "Include Selected Source" & vbLf & "Restore to index scope"
    NMDC_FormatOwnerButton button, RGB(46, 125, 50)

    Set area = ws.Range("K2:M3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Source_Exclude"
    button.OnAction = "NMDC_ExcludeSelectedSource"
    button.TextFrame.Characters.Text = "Exclude Selected Source" & vbLf & "Omit from index scope"
    NMDC_FormatOwnerButton button, RGB(198, 40, 40)
    Exit Sub
Handler:
    ' Source Selection is created by setup; silently skip only in legacy workbooks.
End Sub

Private Sub NMDC_FormatOwnerButton(ByVal button As Shape, ByVal fillColor As Long)
    button.TextFrame.HorizontalAlignment = xlCenter
    button.TextFrame.VerticalAlignment = 3
    button.Fill.ForeColor.RGB = fillColor
    button.Line.ForeColor.RGB = fillColor
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 9
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
End Sub

Private Sub NMDC_ApplyPlainRulesExperience()
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim table As ListObject
    Dim target As Range
    Dim guide As Shape

    Set ws = ThisWorkbook.Worksheets("Rules & Mappings")
    Set table = ws.ListObjects("ClassificationRules")

    Set target = Nothing
    On Error Resume Next
    Set target = table.ListColumns("Match_Type").DataBodyRange
    On Error GoTo Handler
    If Not target Is Nothing Then
        On Error Resume Next
        target.Validation.Delete
        On Error GoTo Handler
        target.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="CONTAINS,EXACT,STARTS_WITH,ENDS_WITH"
        target.Validation.IgnoreBlank = False
        target.Validation.InCellDropdown = True
        target.Validation.ShowInput = True
        target.Validation.InputTitle = "Simple text match"
        target.Validation.InputMessage = "CONTAINS = phrase appears anywhere; EXACT = whole field matches; STARTS_WITH / ENDS_WITH = position-based text match. No REGEX is required."
    End If

    On Error Resume Next
    Set guide = ws.Shapes("NMDC_Rules_Guide")
    On Error GoTo Handler
    If Not guide Is Nothing Then
        guide.TextFrame.Characters.Text = "Normal edit: choose dropdowns and type ordinary words in Match_Words. Use CONTAINS for most rules; EXACT, STARTS_WITH and ENDS_WITH cover the other normal cases. Shipped rules no longer require REGEX."
    End If

    On Error Resume Next
    ws.Range("A4:R4").UnMerge
    ws.Range("A4:R4").Merge
    On Error GoTo Handler
    With ws.Range("A4:R4")
        .Value = "SIMPLE RULE EDITOR. Normal users work with plain text only: choose where to look, choose CONTAINS / EXACT / STARTS WITH / ENDS WITH, type the words, and choose the classification result. Technical columns stay hidden unless explicitly shown."
        .Interior.Color = RGB(255, 247, 219)
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(96, 72, 0)
        .WrapText = True
        .RowHeight = 54
    End With
    Exit Sub
Handler:
    NMDC_LogError "PLAIN_RULES_UX_ERROR", _
        "Excel could not apply the plain-text Rules & Mappings experience.", _
        Err.Number & " - " & Err.Description
End Sub
