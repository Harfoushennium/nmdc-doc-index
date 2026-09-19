Attribute VB_Name = "modNMDC_Startup"
Option Explicit

Public Sub Auto_Open()
    On Error GoTo Handler

    Application.StatusBar = "NMDC Document Index: refreshing dashboard..."
    If NMDC_RunEngine("export-excel") = 0 Then
        If Not NMDC_RefreshExchangeData() Then
            NMDC_LogError "STARTUP_EXCHANGE_ERROR", _
                "The workbook opened, but one or more displayed tables could not be refreshed.", _
                "Open Error Log and press Refresh Dashboard before reviewing or approving an update."
        End If
    End If

    NMDC_ApplyWorkbookGuidance
    Application.StatusBar = False
    Exit Sub

Handler:
    Application.StatusBar = False
    NMDC_LogError "STARTUP_REFRESH_ERROR", _
        "The workbook opened, but the NMDC Index dashboard could not refresh automatically.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ApplyWorkbookGuidance()
    On Error GoTo Handler

    NMDC_SetSheetBanner "Master Documents", _
        "SYSTEM OUTPUT - READ ONLY. One row per extracted document. Filters and sorts are reset automatically whenever data is refreshed so rows cannot become mixed. Use filters and hyperlinks after refresh to review the index."
    NMDC_SetSheetBanner "Revisions", _
        "SYSTEM OUTPUT - READ ONLY. One row per extracted document revision. Filters and sorts are reset on refresh. Do not type corrections directly into this table."
    NMDC_SetSheetBanner "Transactions", _
        "SYSTEM OUTPUT - READ ONLY. One row per extracted event/transaction. Filters and sorts are reset on refresh. Use Source File / Document Link to verify the source record."
    NMDC_SetSheetBanner "Pending Update", _
        "REVIEW ONLY - NO USER INPUT ON THIS SHEET. This table shows what will change if the staged update is approved. Filters and sorts are reset on refresh. Make decisions only in Review Flags, then use Approve / Hold / Reject from Home."
    NMDC_SetSheetBanner "Review Flags", _
        "USER ACTION SHEET. Editable fields are User Decision, User Comment and Resolution Status. Filters and sorts are reset on refresh. Decisions record your review; they do NOT edit the source workbook or silently rewrite extracted data."
    NMDC_SetSheetBanner "User Decisions", _
        "AUDIT TRAIL. This sheet records user/owner decisions and overrides. Treat existing rows as read only; use the workbook actions to create new decisions."
    NMDC_SetSheetBanner "Configuration", _
        "CONFIGURATION. Edit only supported user settings. Package/runtime paths are system settings created by setup. Read each column note before changing a value."
    NMDC_SetSheetBanner "Rules & Mappings", _
        "SIMPLE RULE EDITOR. Use Add Simple Rule, then choose from the dropdowns: Enabled -> Source Family -> Look In (Match Scope) -> Match Method -> Match Words -> Discipline / Category / Subcategory -> Include. Use CONTAINS for most rules; EXACT, STARTS_WITH and ENDS_WITH cover normal alternatives. Legacy technical rule types stay internal. Rules are validated before any scan is staged."
    NMDC_SetSheetBanner "Update History", _
        "AUDIT OUTPUT - READ ONLY. One row per scan/decision. This is the history of staged, approved, held, rejected and review actions."
    NMDC_SetSheetBanner "Error Log", _
        "DIAGNOSTIC OUTPUT - READ ONLY. Plain-English error and recommended action come first; Technical Detail is for troubleshooting/support."

    NMDC_ApplyNamedTableGuidance "Master Documents", "MasterDocuments"
    NMDC_ApplyNamedTableGuidance "Revisions", "RevisionRegister"
    NMDC_ApplyNamedTableGuidance "Transactions", "EventRegister"
    NMDC_ApplyNamedTableGuidance "Pending Update", "PendingUpdate"
    NMDC_ApplyNamedTableGuidance "Review Flags", "ReviewFlags"
    NMDC_ApplyNamedTableGuidance "User Decisions", "UserDecisionLog"
    NMDC_ApplyNamedTableGuidance "Configuration", "Configuration"
    NMDC_ApplyNamedTableGuidance "Rules & Mappings", "ClassificationRules"
    NMDC_ApplyNamedTableGuidance "Update History", "UpdateHistory"
    NMDC_ApplyNamedTableGuidance "Error Log", "ErrorLog"
    NMDC_ConfigureRulesUserExperience
    Exit Sub

Handler:
    NMDC_LogError "WORKBOOK_GUIDANCE_ERROR", _
        "Excel could not apply all worksheet guidance and field notes.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ApplyTableGuidance(ByVal table As ListObject)
    On Error GoTo Handler

    Dim column As ListColumn
    Dim noteText As String
    Dim headerCell As Range

    table.TableStyle = "TableStyleMedium2"
    table.ShowTableStyleRowStripes = False
    table.ShowTableStyleColumnStripes = False
    table.ShowAutoFilter = True

    table.Range.Font.Name = "Aptos"
    table.Range.Font.Size = 10
    table.Range.Font.ColorIndex = xlAutomatic
    table.Range.VerticalAlignment = xlCenter

    With table.HeaderRowRange
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(20, 108, 148)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 34
    End With

    If Not table.DataBodyRange Is Nothing Then
        With table.DataBodyRange
            .Interior.Pattern = xlNone
            .Font.Name = "Aptos"
            .Font.Size = 10
            .Font.ColorIndex = xlAutomatic
            .Font.Bold = False
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Rows.RowHeight = 20
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(225, 230, 235)
            .Borders.Weight = xlHairline
        End With
    End If

    For Each column In table.ListColumns
        Set headerCell = column.Range.Cells(1, 1)
        noteText = NMDC_HeaderHelp(table.Name, CStr(column.Name))
        On Error Resume Next
        headerCell.ClearComments
        headerCell.AddComment noteText
        If Not headerCell.Comment Is Nothing Then
            headerCell.Comment.Visible = False
            headerCell.Comment.Shape.TextFrame.AutoSize = True
        End If
        On Error GoTo Handler
        NMDC_SetColumnWidth column
        NMDC_ApplyColumnAlignment column
    Next column

    NMDC_AdjustTableRows table
    NMDC_StyleUserInputColumns table
    Exit Sub
Handler:
    NMDC_LogError "TABLE_GUIDANCE_ERROR", _
        "Excel could not apply all notes or presentation settings to table " & table.Name & ".", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_ApplyNamedTableGuidance(ByVal sheetName As String, ByVal tableName As String)
    Dim table As ListObject
    On Error Resume Next
    Set table = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    On Error GoTo 0
    If table Is Nothing Then Exit Sub
    NMDC_ApplyTableGuidance table
End Sub

Private Sub NMDC_SetSheetBanner(ByVal sheetName As String, ByVal bannerText As String)
    Dim ws As Worksheet
    Dim bannerAddress As String
    Set ws = ThisWorkbook.Worksheets(sheetName)

    If StrComp(sheetName, "Rules & Mappings", vbTextCompare) = 0 Then
        bannerAddress = "A4:R4"
    Else
        bannerAddress = "A4:O4"
    End If

    On Error Resume Next
    ws.Range(bannerAddress).UnMerge
    ws.Range(bannerAddress).ClearContents
    ws.Range(bannerAddress).Merge
    On Error GoTo 0

    With ws.Range(bannerAddress)
        .Value = bannerText
        .Interior.Color = RGB(255, 247, 219)
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(96, 72, 0)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 54
    End With
End Sub

Private Sub NMDC_SetColumnWidth(ByVal column As ListColumn)
    Dim headerName As String
    headerName = UCase$(Trim$(CStr(column.Name)))

    Select Case headerName
        Case "PLAIN-ENGLISH PROBLEM", "RECOMMENDED USER ACTION", "PLAIN-ENGLISH SUMMARY", "DOCUMENT TITLE", "TECHNICAL DETAIL", "NOTES", "USER NOTE"
            column.Range.ColumnWidth = 42
        Case "SOURCE FILE", "DOCUMENT LINK", "MATCH_WORDS", "EXCLUDE_WORDS", "PATH_QUALIFIER"
            column.Range.ColumnWidth = 34
        Case "USER COMMENT", "EVENT VALUES JSON"
            column.Range.ColumnWidth = 36
        Case "RECORD IDENTITY", "GLOBAL DOCUMENT KEY", "REVISION KEY", "EVENT KEY", "DECISION ID", "RUN ID"
            column.Range.ColumnWidth = 28
        Case "DATE/TIME", "EVENT DATE", "LATEST EVENT DATE", "SOURCE MODIFIED DATE", "CREATED AT", "UPDATED AT"
            column.Range.ColumnWidth = 19
        Case "PROJECT NO.", "DOCUMENT NO.", "COMPANY DOCUMENT NO.", "REVISION", "LATEST REVISION", "EVENT TYPE", "EVENT STATUS", "FLAG CODE", "USER DECISION", "RESOLUTION STATUS", "DECISION TYPE"
            column.Range.ColumnWidth = 22
        Case "DISCIPLINE", "CATEGORY", "SUBCATEGORY"
            column.Range.ColumnWidth = 24
        Case Else
            column.Range.ColumnWidth = 16
    End Select
End Sub

Private Sub NMDC_ApplyColumnAlignment(ByVal column As ListColumn)
    If column.DataBodyRange Is Nothing Then Exit Sub

    Dim headerName As String
    headerName = UCase$(Trim$(CStr(column.Name)))

    Select Case headerName
        Case "PROJECT NO.", "REVISION", "LATEST REVISION", "SOURCE ROW", _
             "DATE/TIME", "EVENT DATE", "LATEST EVENT DATE", "SOURCE MODIFIED DATE", _
             "CREATED AT", "UPDATED AT", "FLAG LEVEL", "FLAG CODE", "EVENT STATUS", _
             "USER DECISION", "RESOLUTION STATUS", "DECISION TYPE", "ENABLED", "PRIORITY", _
             "SOURCE_FAMILY", "MATCH_SCOPE", "MATCH_TYPE", "INCLUDE", "MIN_CONFIDENCE", _
             "STOP_ON_MATCH", "NEW SOURCES", "CHANGED SOURCES", "REMOVED SOURCES", _
             "ADDED RECORDS", "MODIFIED RECORDS", "REMOVED RECORDS", "REVIEW FLAGS", "CONFLICT FLAGS"
            column.DataBodyRange.HorizontalAlignment = xlCenter
        Case Else
            column.DataBodyRange.HorizontalAlignment = xlLeft
    End Select
End Sub

Private Sub NMDC_AdjustTableRows(ByVal table As ListObject)
    If table.DataBodyRange Is Nothing Then Exit Sub

    Dim column As ListColumn
    Dim rowCount As Long
    Dim headerName As String

    rowCount = table.DataBodyRange.Rows.Count
    For Each column In table.ListColumns
        headerName = UCase$(Trim$(CStr(column.Name)))
        Select Case headerName
            Case "DOCUMENT TITLE", "PLAIN-ENGLISH PROBLEM", "RECOMMENDED USER ACTION", _
                 "PLAIN-ENGLISH SUMMARY", "USER COMMENT", "TECHNICAL DETAIL", "NOTES", "USER NOTE"
                If Not column.DataBodyRange Is Nothing Then column.DataBodyRange.WrapText = True
        End Select
    Next column

    If rowCount <= 2500 Then
        table.DataBodyRange.Rows.AutoFit
        Dim oneRow As Range
        For Each oneRow In table.DataBodyRange.Rows
            If oneRow.RowHeight < 20 Then oneRow.RowHeight = 20
            If oneRow.RowHeight > 60 Then oneRow.RowHeight = 60
        Next oneRow
    Else
        table.DataBodyRange.Rows.RowHeight = 20
    End If
End Sub

Private Sub NMDC_StyleUserInputColumns(ByVal table As ListObject)
    Dim target As Range

    If StrComp(table.Name, "ReviewFlags", vbTextCompare) = 0 Then
        Set target = Nothing
        On Error Resume Next
        Set target = table.ListColumns("User Decision").DataBodyRange
        On Error GoTo 0
        If Not target Is Nothing Then
            target.Interior.Pattern = xlNone
            target.Font.ColorIndex = xlAutomatic
        End If

        Set target = Nothing
        On Error Resume Next
        Set target = table.ListColumns("User Comment").DataBodyRange
        On Error GoTo 0
        If Not target Is Nothing Then
            target.Interior.Pattern = xlNone
            target.Font.ColorIndex = xlAutomatic
        End If

        Set target = Nothing
        On Error Resume Next
        Set target = table.ListColumns("Resolution Status").DataBodyRange
        On Error GoTo 0
        If Not target Is Nothing Then
            target.Interior.Pattern = xlNone
            target.Font.ColorIndex = xlAutomatic
        End If
    End If
End Sub

Public Sub NMDC_ConfigureRulesUserExperience()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Set ws = ThisWorkbook.Worksheets("Rules & Mappings")
    Set table = ws.ListObjects("ClassificationRules")

    NMDC_ApplyRulesListValidation table, "Enabled", "YES,NO", "Enable rule", "YES = use this rule; NO = keep it but do not use it."
    NMDC_ApplyRulesListValidation table, "Source_Family", "ANY,METHODS,TECH", "Source family", "ANY = all sources; METHODS = methods registers only; TECH = technical registers only."
    NMDC_ApplyRulesListValidation table, "Match_Scope", "FILE,WORKSHEET,SECTION,HEADER,DOC_NUMBER,TITLE", "Where should Excel look?", "Choose FILE, WORKSHEET, SECTION, HEADER, DOC_NUMBER or TITLE."
    NMDC_ApplyRulesListValidation table, "Match_Type", "CONTAINS,EXACT,STARTS_WITH,ENDS_WITH", "How should it match?", "CONTAINS is recommended. EXACT matches the complete field. STARTS_WITH and ENDS_WITH match text at the beginning or end."
    NMDC_ApplyRulesListValidation table, "Include", "YES,NO", "Keep or exclude?", "YES = classify/include matching data; NO = exclude matching content."
    NMDC_ApplyRulesListValidation table, "Stop_On_Match", "YES,NO", "Stop after match?", "YES is normally safest; later lower-priority rules will not override this match."
    NMDC_ApplyRulesWholeNumberValidation table, "Priority", 1, 999999
    NMDC_ApplyRulesDecimalValidation table, "Min_Confidence", 0, 1

    NMDC_SetRulesAdvancedVisibility table, True
    NMDC_CreateRulesToolbar ws, table
    Exit Sub
Handler:
    NMDC_LogError "RULES_UX_ERROR", _
        "Excel could not apply the simple Rules & Mappings editor.", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_ApplyRulesListValidation(ByVal table As ListObject, ByVal columnName As String, ByVal listValues As String, ByVal titleText As String, ByVal messageText As String)
    Dim target As Range
    On Error Resume Next
    Set target = table.ListColumns(columnName).DataBodyRange
    On Error GoTo 0
    If target Is Nothing Then Exit Sub

    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    target.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listValues
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    target.Validation.ShowInput = True
    target.Validation.InputTitle = titleText
    target.Validation.InputMessage = messageText
End Sub

Private Sub NMDC_ApplyRulesWholeNumberValidation(ByVal table As ListObject, ByVal columnName As String, ByVal minimumValue As Long, ByVal maximumValue As Long)
    Dim target As Range
    On Error Resume Next
    Set target = table.ListColumns(columnName).DataBodyRange
    On Error GoTo 0
    If target Is Nothing Then Exit Sub

    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    target.Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:=CStr(minimumValue), Formula2:=CStr(maximumValue)
    target.Validation.IgnoreBlank = False
    target.Validation.ShowInput = True
    target.Validation.InputTitle = "Rule priority"
    target.Validation.InputMessage = "Lower numbers run first. Add Simple Rule creates a safe next priority automatically."
End Sub

Private Sub NMDC_ApplyRulesDecimalValidation(ByVal table As ListObject, ByVal columnName As String, ByVal minimumValue As Double, ByVal maximumValue As Double)
    Dim target As Range
    On Error Resume Next
    Set target = table.ListColumns(columnName).DataBodyRange
    On Error GoTo 0
    If target Is Nothing Then Exit Sub

    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    target.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:=CStr(minimumValue), Formula2:=CStr(maximumValue)
    target.Validation.IgnoreBlank = True
    target.Validation.ShowInput = True
    target.Validation.InputTitle = "Confidence"
    target.Validation.InputMessage = "Use a value from 0 to 1. This is an advanced setting; 0.9 is the normal default."
End Sub

Private Sub NMDC_CreateRulesToolbar(ByVal ws As Worksheet, ByVal table As ListObject)
    Dim shape As Shape
    Dim area As Range
    Dim button As Shape

    On Error Resume Next
    ws.Shapes("NMDC_Rules_Add").Delete
    ws.Shapes("NMDC_Rules_Advanced").Delete
    ws.Shapes("NMDC_Rules_Guide").Delete
    On Error GoTo 0

    Set area = ws.Range("A2:C3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Rules_Add"
    button.OnAction = "NMDC_AddSimpleRule"
    button.TextFrame.Characters.Text = "Add Simple Rule"
    NMDC_FormatRulesButton button, RGB(46, 125, 50)

    Set area = ws.Range("D2:F3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Rules_Advanced"
    button.OnAction = "NMDC_ToggleRulesAdvancedColumns"
    button.TextFrame.Characters.Text = "Show / Hide Advanced"
    NMDC_FormatRulesButton button, RGB(91, 100, 112)

    Set area = ws.Range("G2:R3")
    Set shape = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    shape.Name = "NMDC_Rules_Guide"
    shape.TextFrame.Characters.Text = "Normal edit: choose dropdowns and type ordinary words in Match_Words. Use CONTAINS for most rules; EXACT, STARTS_WITH and ENDS_WITH cover the normal alternatives. Technical legacy rule types stay outside the normal owner workflow."
    shape.Fill.ForeColor.RGB = RGB(247, 249, 252)
    shape.Line.ForeColor.RGB = RGB(216, 225, 232)
    shape.TextFrame.Characters.Font.Name = "Aptos"
    shape.TextFrame.Characters.Font.Size = 10
    shape.TextFrame.Characters.Font.Color = RGB(31, 41, 55)
    shape.TextFrame.VerticalAlignment = 3
End Sub

Private Sub NMDC_FormatRulesButton(ByVal button As Shape, ByVal fillColor As Long)
    button.TextFrame.HorizontalAlignment = -4108
    button.TextFrame.VerticalAlignment = 3
    button.Fill.ForeColor.RGB = fillColor
    button.Line.ForeColor.RGB = fillColor
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 10
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
End Sub

Private Sub NMDC_SetRulesAdvancedVisibility(ByVal table As ListObject, ByVal hideAdvanced As Boolean)
    Dim names As Variant
    Dim item As Variant
    names = Array("Rule_ID", "Priority", "Exclude_Words", "Path_Qualifier", _
                  "Requires_Discipline", "Requires_Category", "Min_Confidence", "Stop_On_Match")

    For Each item In names
        On Error Resume Next
        table.ListColumns(CStr(item)).Range.EntireColumn.Hidden = hideAdvanced
        On Error GoTo 0
    Next item
End Sub

Public Sub NMDC_ToggleRulesAdvancedColumns()
    On Error GoTo Handler

    Dim table As ListObject
    Dim currentlyHidden As Boolean
    Set table = ThisWorkbook.Worksheets("Rules & Mappings").ListObjects("ClassificationRules")

    currentlyHidden = table.ListColumns("Rule_ID").Range.EntireColumn.Hidden
    NMDC_SetRulesAdvancedVisibility table, Not currentlyHidden
    Exit Sub
Handler:
    MsgBox "Excel could not change the Rules & Mappings view." & vbCrLf & Err.Description, _
           vbExclamation, "NMDC Document Index"
End Sub

Public Sub NMDC_AddSimpleRule()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim newRow As ListRow
    Dim newId As String
    Dim newPriority As Long

    Set ws = ThisWorkbook.Worksheets("Rules & Mappings")
    Set table = ws.ListObjects("ClassificationRules")
    Set newRow = table.ListRows.Add

    newId = NMDC_NextUserRuleId(table)
    newPriority = NMDC_NextRulePriority(table)

    NMDC_SetRuleCell table, newRow, "Rule_ID", newId
    NMDC_SetRuleCell table, newRow, "Enabled", "YES"
    NMDC_SetRuleCell table, newRow, "Priority", newPriority
    NMDC_SetRuleCell table, newRow, "Source_Family", "ANY"
    NMDC_SetRuleCell table, newRow, "Match_Scope", "WORKSHEET"
    NMDC_SetRuleCell table, newRow, "Match_Type", "CONTAINS"
    NMDC_SetRuleCell table, newRow, "Match_Words", ""
    NMDC_SetRuleCell table, newRow, "Discipline", ""
    NMDC_SetRuleCell table, newRow, "Category", ""
    NMDC_SetRuleCell table, newRow, "Subcategory", ""
    NMDC_SetRuleCell table, newRow, "Include", "YES"
    NMDC_SetRuleCell table, newRow, "Min_Confidence", 0.9
    NMDC_SetRuleCell table, newRow, "Stop_On_Match", "YES"
    NMDC_SetRuleCell table, newRow, "Notes", "User rule - describe the purpose in plain English"

    NMDC_ConfigureRulesUserExperience
    ws.Activate
    newRow.Range.Cells(1, table.ListColumns("Match_Words").Index).Select
    MsgBox "A new simple rule was added." & vbCrLf & vbCrLf & _
           "1. Choose where to look and how to match." & vbCrLf & _
           "2. Type the word or phrase to match." & vbCrLf & _
           "3. Enter the Discipline, Category and Subcategory result." & vbCrLf & _
           "4. Run Update Changed Files or Full Rescan; the rules will be validated before staging.", _
           vbInformation, "NMDC Document Index"
    Exit Sub
Handler:
    MsgBox "Excel could not add a new rule." & vbCrLf & Err.Description, vbExclamation, "NMDC Document Index"
End Sub

Private Sub NMDC_SetRuleCell(ByVal table As ListObject, ByVal row As ListRow, ByVal columnName As String, ByVal value As Variant)
    row.Range.Cells(1, table.ListColumns(columnName).Index).Value = value
End Sub

Private Function NMDC_NextUserRuleId(ByVal table As ListObject) As String
    Dim index As Long
    Dim candidate As String
    index = 1
    Do
        candidate = "USR" & Format$(index, "000")
        If Not NMDC_RuleIdExists(table, candidate) Then
            NMDC_NextUserRuleId = candidate
            Exit Function
        End If
        index = index + 1
    Loop
End Function

Private Function NMDC_RuleIdExists(ByVal table As ListObject, ByVal candidate As String) As Boolean
    Dim row As ListRow
    Dim colIndex As Long
    colIndex = table.ListColumns("Rule_ID").Index
    For Each row In table.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, colIndex).Value)), candidate, vbTextCompare) = 0 Then
            NMDC_RuleIdExists = True
            Exit Function
        End If
    Next row
End Function

Private Function NMDC_NextRulePriority(ByVal table As ListObject) As Long
    Dim row As ListRow
    Dim colIndex As Long
    Dim value As Variant
    Dim maximum As Long

    colIndex = table.ListColumns("Priority").Index
    maximum = 0
    For Each row In table.ListRows
        value = row.Range.Cells(1, colIndex).Value
        If IsNumeric(value) Then
            If CLng(value) > maximum Then maximum = CLng(value)
        End If
    Next row
    NMDC_NextRulePriority = maximum + 10
End Function

Private Function NMDC_HeaderHelp(ByVal tableName As String, ByVal headerName As String) As String
    Dim h As String
    Dim isRules As Boolean
    h = UCase$(Trim$(headerName))
    isRules = (StrComp(tableName, "ClassificationRules", vbTextCompare) = 0)

    Select Case h
        Case "FLAG LEVEL"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Overall data-quality status for the row. OK = extracted normally; REVIEW = user attention is requested; CONFLICT = blocking/integrity issue."
        Case "FLAG CODE"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Machine-readable reason for the review flag. Use it together with Plain-English Problem and Recommended User Action."
        Case "PLAIN-ENGLISH PROBLEM"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Human-readable explanation of why this item was flagged."
        Case "RECOMMENDED USER ACTION"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Suggested next step. Follow this guidance, then record your choice in User Decision."
        Case "PROJECT NO."
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Project number assigned to the extracted record. If wrong, do not type over it; use Flag Wrong Data / Review Flags."
        Case "SOURCE FAMILY"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. High-level source group such as METHODS or TECH, derived from the source location/rules."
        Case "DISCIPLINE"
            If isRules Then
                NMDC_HeaderHelp = "USER INPUT. Classification discipline to assign when this rule matches, for example MARINE OPERATIONS or OFFSHORE INSTALLATION."
            Else
                NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Classification discipline assigned by the approved Rules & Mappings."
            End If
        Case "CATEGORY"
            If isRules Then
                NMDC_HeaderHelp = "USER INPUT. Main classification to assign when this rule matches, for example DRAWING, PROCEDURE or DOCUMENT."
            Else
                NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Main document classification assigned by the approved rules."
            End If
        Case "SUBCATEGORY"
            If isRules Then
                NMDC_HeaderHelp = "USER INPUT. More detailed classification to assign when this rule matches, for example ANCHOR PATTERN or INSTALLATION PROCEDURE."
            Else
                NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Detailed document classification assigned by the approved rules."
            End If
        Case "DOCUMENT NO."
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. NMDC/document identifier extracted from the source register. Preserved as text to avoid changing leading zeroes or formatting."
        Case "DOCUMENT TITLE"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Document description/title extracted from the source register."
        Case "COMPANY DOCUMENT NO."
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Client/company document number where available in the source register."
        Case "LATEST REVISION"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Latest revision identified for the document from the extracted revision records."
        Case "REVISION"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Revision value exactly as extracted. Kept as text so values such as 00 are not changed by Excel."
        Case "IS LATEST REVISION"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. 1/YES means this row is the latest known revision for that document; otherwise it is historical."
        Case "REVISION KEY"
            NMDC_HeaderHelp = "TECHNICAL OUTPUT - READ ONLY. Stable internal identity used to group and compare one document revision."
        Case "LATEST EVENT DATE"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Date of the latest extracted transaction/event for the latest revision."
        Case "LATEST EVENT STATUS"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Status associated with the latest extracted event."
        Case "EVENT TYPE"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Type/channel of the extracted register event or transaction."
        Case "EVENT DATE"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Date of the extracted event/transaction."
        Case "EVENT REFERENCE"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Reference/transmittal identifier associated with the event."
        Case "EVENT STATUS"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Status/code recorded for the event in the source register."
        Case "EVENT VALUES JSON"
            NMDC_HeaderHelp = "TECHNICAL OUTPUT - READ ONLY. Lossless copy of additional event fields not promoted to dedicated columns. Mainly for audit/troubleshooting."
        Case "IS LATEST EVENT"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. 1/YES means this is the latest known event for that revision."
        Case "EVENT KEY"
            NMDC_HeaderHelp = "TECHNICAL OUTPUT - READ ONLY. Stable internal identity for an extracted event; used for reconciliation and review decisions."
        Case "DOCUMENT LINK"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - CLICKABLE. Opens the source document/hyperlink where one was available in the register."
        Case "SOURCE FILE", "RELATIVE PATH", "RELATIVE_PATH"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - CLICKABLE. Source workbook from which this row was extracted. Click to open and verify the original data."
        Case "SOURCE SHEET", "WORKSHEET"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Worksheet/tab inside the source workbook where the record came from."
        Case "SOURCE ROW"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Original source worksheet row number used for traceability."
        Case "SOURCE CELL", "SOURCE ROW/CELL"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Original source cell/row reference used for traceability and troubleshooting."
        Case "GLOBAL DOCUMENT KEY"
            NMDC_HeaderHelp = "TECHNICAL OUTPUT - READ ONLY. Internal canonical identity used to group the same document across source records."
        Case "CHANGE TYPE"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. ADDED = new staged record; MODIFIED = differs from approved; REMOVED = would disappear after approval; UNCHANGED = carried forward unchanged."
        Case "RECORD IDENTITY"
            NMDC_HeaderHelp = "TECHNICAL OUTPUT - READ ONLY. Stable identity used to compare approved and staged records."
        Case "PLAIN-ENGLISH SUMMARY"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. Explains what this staged row means if the update is approved."
        Case "REVIEW REQUIRED"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. YES means the change deserves user attention before approval. Review decisions are entered on Review Flags, not Pending Update."
        Case "USER DECISION"
            NMDC_HeaderHelp = "USER INPUT. ACKNOWLEDGED = I reviewed the flag; extraction is unchanged. NO ACTION REQUIRED = accept current extraction for this flag. NEEDS SOURCE CORRECTION = source workbook must be corrected and rescanned; no automatic edit occurs. NEEDS PARSER/MAPPING FIX = parser/rule needs correction and a new scan; no automatic edit occurs. HOLD FOR REVIEW = leave unresolved; do not approve yet. Saving a decision records it; it does not silently rewrite extracted data."
        Case "USER COMMENT"
            NMDC_HeaderHelp = "USER INPUT - FREE TEXT. Explain your decision, expected value, or what needs correction. Required in practice for source/parser corrections so the follow-up is unambiguous."
        Case "RESOLUTION STATUS"
            NMDC_HeaderHelp = "USER INPUT. OPEN = not resolved; ACKNOWLEDGED = reviewed but not necessarily fixed; RESOLVED = issue is considered closed; DEFERRED = intentionally postponed. This status records review state; it does not modify the source file."
        Case "DATE/TIME"
            NMDC_HeaderHelp = "SYSTEM/AUDIT OUTPUT - READ ONLY. Date and time when the logged event, decision or error occurred."
        Case "EVENT"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Type of history event recorded by the engine."
        Case "RUN ID"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Unique identifier of the scan/staged/approved run."
        Case "MODE"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. INCREMENTAL processes detected changes; FULL_RESCAN rebuilds from all selected sources."
        Case "DECISION"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Lifecycle decision such as STAGED, APPROVED, HOLD or REJECTED."
        Case "STATUS"
            NMDC_HeaderHelp = "SYSTEM/AUDIT OUTPUT unless this sheet specifically identifies it as editable. Current state of the item or run."
        Case "NEW SOURCES", "CHANGED SOURCES", "REMOVED SOURCES", "ADDED RECORDS", "MODIFIED RECORDS", "REMOVED RECORDS", "REVIEW FLAGS", "CONFLICT FLAGS"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Count recorded for this run/decision."
        Case "NOTE"
            NMDC_HeaderHelp = "AUDIT OUTPUT. Optional note captured with a hold/reject/review action."
        Case "SEVERITY"
            NMDC_HeaderHelp = "DIAGNOSTIC OUTPUT - READ ONLY. Importance level of the logged issue."
        Case "ACTION"
            NMDC_HeaderHelp = "DIAGNOSTIC OUTPUT - READ ONLY. Operation/error code that created this log entry."
        Case "PLAIN-ENGLISH ERROR"
            NMDC_HeaderHelp = "DIAGNOSTIC OUTPUT - READ ONLY. User-facing explanation of what went wrong."
        Case "RECOMMENDED ACTION"
            NMDC_HeaderHelp = "DIAGNOSTIC OUTPUT - READ ONLY. Recommended recovery or support step."
        Case "TECHNICAL DETAIL"
            NMDC_HeaderHelp = "DIAGNOSTIC OUTPUT - READ ONLY. Technical evidence for troubleshooting. Normally not required for routine user decisions."
        Case "SETTING"
            NMDC_HeaderHelp = "CONFIGURATION KEY. Identifies the setting. Do not rename system settings because VBA/engine look them up by this exact name."
        Case "CURRENT VALUE"
            NMDC_HeaderHelp = "USER/SYSTEM CONFIGURATION VALUE. Change only supported user settings. Setup-managed engine/runtime/configuration paths should normally be left unchanged."
        Case "RULE_ID"
            NMDC_HeaderHelp = "ADVANCED RULE IDENTIFIER. Add Simple Rule creates this automatically. Keep it unique and stable because the index uses it for audit history."
        Case "ENABLED"
            NMDC_HeaderHelp = "USER INPUT - DROPDOWN. YES = use the rule. NO = keep the rule in the sheet but ignore it during classification."
        Case "PRIORITY"
            NMDC_HeaderHelp = "ADVANCED USER INPUT. Lower numbers run first. Add Simple Rule assigns a safe next priority automatically."
        Case "SOURCE_FAMILY"
            If isRules Then
                NMDC_HeaderHelp = "USER INPUT - DROPDOWN. ANY = all registers; METHODS = methods registers only; TECH = technical/document registers only."
            Else
                NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY. High-level source group such as METHODS or TECH."
            End If
        Case "MATCH_SCOPE"
            NMDC_HeaderHelp = "USER INPUT - DROPDOWN. Choose where to look: FILE = file/path name; WORKSHEET = sheet/tab name; SECTION = section name; HEADER = table headers; DOC_NUMBER = document number; TITLE = document title."
        Case "MATCH_TYPE"
            NMDC_HeaderHelp = "USER INPUT - DROPDOWN. CONTAINS is recommended. EXACT matches the whole field. STARTS_WITH and ENDS_WITH match text at the beginning or end."
        Case "MATCH_WORDS"
            NMDC_HeaderHelp = "USER INPUT. Type ordinary words or a phrase exactly as you expect to see it. The selected Match Type controls whether the text may appear anywhere, must match the whole field, or must appear at the beginning/end."
        Case "EXCLUDE_WORDS"
            NMDC_HeaderHelp = "ADVANCED USER INPUT. Optional pattern that prevents this rule from matching. Leave blank unless you need an explicit exception."
        Case "PATH_QUALIFIER"
            NMDC_HeaderHelp = "ADVANCED USER INPUT. Optional source-path restriction. Leave blank for normal rules."
        Case "REQUIRES_DISCIPLINE"
            NMDC_HeaderHelp = "ADVANCED USER INPUT. Optional existing discipline that must already be assigned before this rule is allowed to refine a record."
        Case "REQUIRES_CATEGORY"
            NMDC_HeaderHelp = "ADVANCED USER INPUT. Optional existing category that must already be assigned before this rule is allowed to refine a record."
        Case "INCLUDE"
            NMDC_HeaderHelp = "USER INPUT - DROPDOWN. YES = classify/include matching content. NO = intentionally exclude matching content from the index."
        Case "MIN_CONFIDENCE"
            NMDC_HeaderHelp = "ADVANCED INTERNAL INPUT. Confidence threshold retained for backward compatibility. Normal owner-created text rules do not need to change it."
        Case "STOP_ON_MATCH"
            NMDC_HeaderHelp = "ADVANCED USER INPUT - YES/NO. YES normally prevents lower-priority rules from changing the result after this rule matches."
        Case "NOTES"
            NMDC_HeaderHelp = "USER INPUT / DOCUMENTATION. Describe the rule purpose in plain English so another user can understand why it exists."
        Case "DECISION ID"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Unique identity of a recorded user decision."
        Case "USER"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Excel/user name associated with the decision."
        Case "DECISION TYPE"
            NMDC_HeaderHelp = "USER/AUDIT FIELD. Type of recorded owner decision such as APPROVE, HOLD, REJECT, OVERRIDE or CONFIGURATION CHANGE."
        Case "PREVIOUS VALUE"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Value before an approved user decision/override."
        Case "APPROVED VALUE"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Value explicitly accepted through the decision workflow."
        Case "CONFIGURATION VERSION"
            NMDC_HeaderHelp = "AUDIT OUTPUT - READ ONLY. Configuration/rule version associated with the decision or extraction."
        Case Else
            NMDC_HeaderHelp = "SYSTEM FIELD. Generated or maintained by the NMDC Document Index. Read this field for filtering/audit; do not overwrite extracted output directly. Use the supported review/configuration workflow for corrections."
    End Select
End Function
