Attribute VB_Name = "modNMDC_OwnerUX"
Option Explicit

Public Sub NMDC_ApplyOwnerUX()
    On Error GoTo Handler
    NMDC_ApplyHomeButtonGuides
    NMDC_ApplyHomeWorkflowGuide
    NMDC_ApplyPendingUpdateDecisionGuide
    NMDC_ApplyPlainRulesExperience
    Exit Sub
Handler:
    NMDC_LogError "OWNER_UX_ERROR", _
        "Excel could not apply all owner guidance controls.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_OpenSourceSelection()
    ' Backward-compatible action: the action now opens File Review.
    On Error GoTo Handler
    If Not NMDC_ReviewDraftStoreHealthy() Or Not NMDC_ReviewDraftCaptureCurrentTables() Then Exit Sub
    If NMDC_RunEngine("export-excel") <> 0 Then
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    If Not NMDC_RefreshReviewDataFast() Then
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    NMDC_GoToSheet "File Review"
    NMDC_RebuildPendingSourcePanel
    NMDC_GoToSourceChoices
    Exit Sub
Handler:
    NMDC_LogError "SOURCE_SELECTION_OPEN_ERROR", _
        "Excel could not open File Review.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_GoToSourceChoices()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim sourceTable As ListObject

    Set ws = ThisWorkbook.Worksheets("File Review")
    Set sourceTable = Nothing
    On Error Resume Next
    Set sourceTable = ws.ListObjects("FileReview")
    On Error GoTo Handler
    If sourceTable Is Nothing Then NMDC_RebuildPendingSourcePanel
    Set sourceTable = ws.ListObjects("FileReview")
    If sourceTable Is Nothing Then Err.Raise vbObjectError + 974, "NMDC Source Selection", "Source choices table is unavailable."
    Application.Goto Reference:=sourceTable.HeaderRowRange, Scroll:=True
    ActiveWindow.ScrollRow = Application.Max(1, sourceTable.HeaderRowRange.Row - 2)
    Exit Sub

Handler:
    NMDC_LogError "SOURCE_SELECTION_OPEN_ERROR", _
        "Excel could not navigate to file choices on File Review.", Err.Number & " - " & Err.Description
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
    Dim table As ListObject, rowIndex As Long
    If ActiveSheet.Name <> "File Review" Then NMDC_GoToSourceChoices: Exit Sub
    Set table = ActiveSheet.ListObjects("FileReview")
    If table.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Selection, table.DataBodyRange) Is Nothing Then Exit Sub
    rowIndex = Selection.Cells(1, 1).Row - table.DataBodyRange.Row + 1
    table.ListColumns("Accept Update?").DataBodyRange.Cells(rowIndex, 1).Value2 = includeSource
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
    ' Compatibility entry point. There is no separate owner-facing source sheet now.
    On Error GoTo Handler
    NMDC_RefreshSourceSelectionTable = NMDC_RebuildPendingSourcePanel()
    Exit Function
Handler:
    NMDC_LogError "SOURCE_SELECTION_REFRESH_ERROR", _
        "Excel could not refresh File Review.", Err.Number & " - " & Err.Description
    NMDC_RefreshSourceSelectionTable = False
End Function

Private Sub NMDC_ApplyHomeButtonGuides()
    Dim titles As Variant
    Dim guides As Variant
    Dim index As Long
    Dim shape As Shape

    titles = Array( _
        "Update Changed Files", "Full Rescan / Rebuild All", "Review File Updates", _
        "Review File Choices", "Hold Update", "Reject Update", _
        "Select Data Folder", "Review Flags", "Save Review Decisions", _
        "Configuration", "Rules & Mappings", "View Log", _
        "Flag Wrong Data", "Report Requirement / Problem", "Refresh Dashboard", _
        "Undo Last Approval", "Reset All Records", "Help")

    guides = Array( _
        "NORMAL USE - scan new/changed files only", _
        "HEAVY - rebuild every source from the local cache", _
        "Review affected files; optional change details", _
        "Choose files to accept; unchecked keeps approved data", _
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

    NMDC_SafeMergeAndSet ws, "A30:L30", "QUICK WORKFLOW - WHAT TO DO"
    With ws.Range("A30:L30")
        .Interior.Color = RGB(20, 108, 148)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 22
    End With

    NMDC_SetGuideRow ws, 31, "1", "SELECT SOURCE", "Select Data Folder once. Application/runtime stays local; source files may remain in OneDrive."
    NMDC_SetGuideRow ws, 32, "2", "SCAN", "Use Update Changed Files normally. Full Rescan is only for deliberate rebuilds or troubleshooting."
    NMDC_SetGuideRow ws, 33, "3", "REVIEW", "Open File Review to review the affected files. Change Details is optional."
    NMDC_SetGuideRow ws, 34, "4", "CHOOSE FILES", "Check files to accept their update. Uncheck to retain approved data for this proposal. Save File Choices once."
    NMDC_SetGuideRow ws, 35, "5", "DECIDE", "Approve = accept the whole remaining proposal; Hold = postpone; Reject = discard the staged proposal."
    NMDC_SetGuideRow ws, 36, "6", "EXCEPTIONS", "Review Flags is only for genuine extraction/data exceptions. The validated current DATA should produce zero extraction flags."
    NMDC_SetGuideRow ws, 37, "TIP", "LIVE FILTER", "Use the REV03 search box on the sheet. Press Ctrl+Shift+F only to choose the target table header, then type directly in TxtBox_Search. Spaces or + = AND; -word = EXCLUDE; quoted text = exact. RESET SEARCH clears the filter."
    NMDC_SetGuideRow ws, 38, "TIP", "RULES", "Rules & Mappings uses normal text matching: CONTAINS, EXACT, STARTS WITH or ENDS WITH."
End Sub

Private Sub NMDC_SetGuideRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal stepText As String, ByVal titleText As String, ByVal detailText As String)
    NMDC_SafeMergeAndSet ws, "A" & rowNo & ":B" & rowNo, stepText
    NMDC_SafeMergeAndSet ws, "C" & rowNo & ":D" & rowNo, titleText
    NMDC_SafeMergeAndSet ws, "E" & rowNo & ":L" & rowNo, detailText

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
    Dim sourceButton As Shape
    Dim guide As Shape
    Dim anchor As Range
    Const SOURCE_BUTTON_WIDTH As Double = 150
    Const GUIDE_GAP As Double = 8
    Const GUIDE_WIDTH As Double = 600
    Set ws = ThisWorkbook.Worksheets("Change Details")

    On Error Resume Next
    If CStr(ws.Range("A2").Value2) = "No staged update is included in this review-only baseline. Use the macro-enabled package to create a fresh staged comparison." Then
        ws.Range("A2:I2").UnMerge
        ws.Range("A2:I2").ClearContents
    End If
    ws.Shapes("NMDC_Pending_SourceNavigation").Delete
    ws.Shapes("NMDC_Pending_SourceGuide").Delete
    On Error GoTo 0

    Set anchor = ws.Range("A2")
    Set sourceButton = ws.Shapes.AddShape(5, anchor.Left, anchor.Top, SOURCE_BUTTON_WIDTH, 26)
    sourceButton.Name = "NMDC_Pending_SourceNavigation"
    sourceButton.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!NMDC_GoToSourceChoices"
    sourceButton.TextFrame.Characters.Text = "FILE REVIEW"
    sourceButton.AlternativeText = "Open File Review to choose the files to update."
    With sourceButton
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = 3
        .Fill.ForeColor.RGB = RGB(20, 108, 148)
        .Line.ForeColor.RGB = RGB(20, 108, 148)
        .TextFrame.Characters.Font.Name = "Aptos"
        .TextFrame.Characters.Font.Size = 8.5
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Placement = xlFreeFloating
    End With

    Set guide = ws.Shapes.AddShape(1, anchor.Left + SOURCE_BUTTON_WIDTH + GUIDE_GAP, anchor.Top, GUIDE_WIDTH, 26)
    guide.Name = "NMDC_Pending_SourceGuide"
    guide.TextFrame.Characters.Text = _
        "Optional details. Review and accept files on File Review."
    With guide
        .Fill.ForeColor.RGB = RGB(255, 247, 219)
        .Line.ForeColor.RGB = RGB(255, 247, 219)
        .TextFrame.HorizontalAlignment = xlLeft
        .TextFrame.VerticalAlignment = 3
        .TextFrame.MarginLeft = 7
        .TextFrame.MarginRight = 7
        .TextFrame.MarginTop = 2
        .TextFrame.MarginBottom = 2
        .TextFrame.Characters.Font.Name = "Aptos"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(96, 72, 0)
        .Placement = xlFreeFloating
    End With

    NMDC_SafeMergeAndSet ws, "A4:I4", _
        "OPTIONAL CHANGE DETAILS. Review and accept the affected files on File Review. Unchecked files retain their approved data."
    With ws.Range("A4:I4")
        .Interior.Color = RGB(255, 247, 219)
        .Font.Name = "Aptos"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(96, 72, 0)
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .RowHeight = 46
    End With

    On Error Resume Next
    ws.Shapes("NMDC_Pending_Exclude").Delete
    ws.Shapes("NMDC_Pending_SourceList").Delete
    On Error GoTo 0
End Sub

Private Sub NMDC_ApplyPlainRulesExperience()
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim table As ListObject
    Dim target As Range
    Dim guide As Shape

    If Not NMDC_EnsureClassificationRulesTable() Then Exit Sub
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
        target.Validation.InputMessage = "CONTAINS = phrase appears anywhere; EXACT = whole field matches; STARTS_WITH / ENDS_WITH = position-based text match."
    End If

    On Error Resume Next
    Set guide = ws.Shapes("NMDC_Rules_Guide")
    On Error GoTo Handler
    If Not guide Is Nothing Then
        guide.TextFrame.Characters.Text = "Normal edit: choose dropdowns and type ordinary words in Match_Words. CONTAINS is the normal choice; EXACT, STARTS_WITH and ENDS_WITH cover the other simple cases."
    End If

    NMDC_SafeMergeAndSet ws, "A4:R4", _
        "SIMPLE RULE EDITOR. Choose where to look, choose CONTAINS / EXACT / STARTS WITH / ENDS WITH, type ordinary keywords, and choose the classification result. Technical columns stay hidden unless explicitly shown."
    With ws.Range("A4:R4")
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

Private Sub NMDC_SafeMergeAndSet(ByVal ws As Worksheet, ByVal addressText As String, ByVal valueText As String)
    Dim target As Range
    Set target = ws.Range(addressText)

    On Error Resume Next
    target.UnMerge
    target.ClearContents
    On Error GoTo 0
    target.Merge

    target.Cells(1, 1).Value = valueText
End Sub
