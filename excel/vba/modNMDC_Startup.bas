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
        "SYSTEM OUTPUT - READ ONLY. One row per extracted document. Use filters and hyperlinks to review the index. Report corrections through Review Flags or Flag Wrong Data."
    NMDC_SetSheetBanner "Revisions", _
        "SYSTEM OUTPUT - READ ONLY. One row per extracted document revision. Do not type corrections directly into this table."
    NMDC_SetSheetBanner "Transactions", _
        "SYSTEM OUTPUT - READ ONLY. One row per extracted event/transaction. Use Source File / Document Link to verify the source record."
    NMDC_SetSheetBanner "Pending Update", _
        "REVIEW ONLY - NO USER INPUT ON THIS SHEET. This table shows what will change if the staged update is approved. Make review decisions only in Review Flags, then use Approve / Hold / Reject from Home."
    NMDC_SetSheetBanner "Review Flags", _
        "USER ACTION SHEET. Editable fields are User Decision, User Comment and Resolution Status. Decisions record your review; they do NOT edit the source workbook or silently rewrite extracted data. See the notes on each column header before choosing a decision."
    NMDC_SetSheetBanner "User Decisions", _
        "AUDIT TRAIL. This sheet records user/owner decisions and overrides. Treat existing rows as read only; use the workbook actions to create new decisions."
    NMDC_SetSheetBanner "Configuration", _
        "CONFIGURATION. Edit only supported user settings. Package/runtime paths are system settings created by setup. Read each column note before changing a value."
    NMDC_SetSheetBanner "Rules & Mappings", _
        "USER CONFIGURATION. Rules control classification/extraction scope. Use dropdowns where provided, read header notes, then save/stage the rules for review."
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
    table.Range.Font.Name = "Aptos"
    table.Range.Font.Size = 10
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
        .RowHeight = 32
    End With

    If Not table.DataBodyRange Is Nothing Then
        table.DataBodyRange.Rows.RowHeight = 19
        table.DataBodyRange.VerticalAlignment = xlCenter
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
    Next column

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
    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error Resume Next
    ws.Range("A4:O4").UnMerge
    ws.Range("A4:O4").Merge
    On Error GoTo 0

    With ws.Range("A4:O4")
        .Value = bannerText
        .Interior.Color = RGB(255, 247, 219)
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(96, 72, 0)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 44
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
        Case Else
            column.Range.ColumnWidth = 16
    End Select
End Sub

Private Sub NMDC_StyleUserInputColumns(ByVal table As ListObject)
    Dim target As Range

    If StrComp(table.Name, "ReviewFlags", vbTextCompare) = 0 Then
        Set target = Nothing
        On Error Resume Next
        Set target = table.ListColumns("User Decision").DataBodyRange
        On Error GoTo 0
        If Not target Is Nothing Then target.Interior.Color = RGB(234, 244, 251)

        Set target = Nothing
        On Error Resume Next
        Set target = table.ListColumns("User Comment").DataBodyRange
        On Error GoTo 0
        If Not target Is Nothing Then target.Interior.Color = RGB(255, 247, 219)

        Set target = Nothing
        On Error Resume Next
        Set target = table.ListColumns("Resolution Status").DataBodyRange
        On Error GoTo 0
        If Not target Is Nothing Then target.Interior.Color = RGB(234, 246, 236)
    End If
End Sub

Private Function NMDC_HeaderHelp(ByVal tableName As String, ByVal headerName As String) As String
    Dim h As String
    h = UCase$(Trim$(headerName))

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
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY in index tables. Classification discipline assigned by the approved Rules & Mappings."
        Case "CATEGORY"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY in index tables. Main document classification assigned by the approved rules."
        Case "SUBCATEGORY"
            NMDC_HeaderHelp = "SYSTEM OUTPUT - READ ONLY in index tables. Detailed document classification assigned by the approved rules."
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
            NMDC_HeaderHelp = "RULE IDENTIFIER. Keep unique and stable. Used to audit which classification rule matched a source."
        Case "ENABLED"
            NMDC_HeaderHelp = "USER INPUT - DROPDOWN YES/NO. YES allows the rule to participate; NO keeps the rule but disables it."
        Case "PRIORITY"
            NMDC_HeaderHelp = "USER INPUT. Numeric order used when multiple rules may match; lower number is evaluated earlier."
        Case "SOURCE_FAMILY"
            NMDC_HeaderHelp = "USER INPUT. Restricts the rule to a source family (for example METHODS/TECH) or ANY."
        Case "MATCH_SCOPE"
            NMDC_HeaderHelp = "USER INPUT. Defines whether matching applies to file name/path, worksheet or another supported scope."
        Case "MATCH_TYPE"
            NMDC_HeaderHelp = "USER INPUT. Matching method such as CONTAINS or REGEX. Use only supported values because invalid rules are rejected."
        Case "MATCH_WORDS"
            NMDC_HeaderHelp = "USER INPUT - FREE TEXT. Keyword/pattern used by the rule. For REGEX rules this is a regular expression."
        Case "EXCLUDE_WORDS"
            NMDC_HeaderHelp = "USER INPUT - FREE TEXT. Optional words/patterns that prevent this rule from matching."
        Case "PATH_QUALIFIER"
            NMDC_HeaderHelp = "USER INPUT - FREE TEXT. Optional path condition used to narrow where the rule applies."
        Case "REQUIRES_DISCIPLINE"
            NMDC_HeaderHelp = "USER INPUT. Optional prerequisite discipline required before this rule can match."
        Case "REQUIRES_CATEGORY"
            NMDC_HeaderHelp = "USER INPUT. Optional prerequisite category required before this rule can match."
        Case "INCLUDE"
            NMDC_HeaderHelp = "USER INPUT - DROPDOWN YES/NO. YES includes matching content in extraction; NO excludes matching content from the index."
        Case "MIN_CONFIDENCE"
            NMDC_HeaderHelp = "USER INPUT. Minimum confidence threshold required for the rule result to be accepted."
        Case "STOP_ON_MATCH"
            NMDC_HeaderHelp = "USER INPUT - YES/NO. YES stops evaluating lower-priority rules once this rule matches."
        Case "NOTES"
            NMDC_HeaderHelp = "USER INPUT / DOCUMENTATION. Plain-English explanation of the rule purpose for future users and reviewers."
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
