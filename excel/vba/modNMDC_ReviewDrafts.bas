Attribute VB_Name = "modNMDC_ReviewDrafts"
Option Explicit

Private Const DRAFT_FILE As String = "owner_review_drafts.csv"
Private Const DRAFT_HEADERS As String = "Run ID,Section,Identity,Field,Value"
Private Const SOURCE_SECTION As String = "SOURCE"
Private Const FLAGS_SECTION As String = "FLAGS"
Private Const STATE_SECTION As String = "STATE"
Private mReviewDraftStoreFaulted As Boolean

Public Function NMDC_ReviewDraftStoreHealthy() As Boolean
    On Error GoTo Handler
    Dim drafts As Object
    If mReviewDraftStoreFaulted Then GoTo Handler
    Set drafts = NMDC_ReviewDraftLoad()
    NMDC_ReviewDraftStoreHealthy = True
    Exit Function
Handler:
    NMDC_ReviewDraftStoreHealthy = False
    NMDC_ReviewDraftReportStoreError "Excel could not read the saved review draft store. Refresh and approval are blocked.", _
        Err.Number & " - " & Err.Description
End Function

' Defensive seam for edits made before event hooks were installed or while
' Application.EnableEvents was disabled by another Excel action. Compare only
' editable cells against the exchange baseline for the same displayed run.
Public Function NMDC_ReviewDraftCaptureCurrentTables() As Boolean
    On Error GoTo Handler

    Dim runId As String
    Dim dashboardHeaders As Variant, dashboardData As Variant
    Dim sourceHeaders As Variant, sourceData As Variant
    Dim flagHeaders As Variant, flagData As Variant
    Dim dashboardRows As Long, dashboardCols As Long
    Dim sourceRows As Long, sourceCols As Long
    Dim flagRows As Long, flagCols As Long
    Dim drafts As Object
    Dim ws As Worksheet
    Dim table As ListObject
    Dim rowIndex As Long, csvRow As Long
    Dim sourceKey As String, flagKey As String
    Dim fieldName As Variant
    Dim currentValue As String, baselineValue As String
    Dim changed As Boolean

    runId = NMDC_DisplayedPendingRunId()
    If Len(runId) = 0 Then NMDC_ReviewDraftCaptureCurrentTables = True: Exit Function
    If Not NMDC_ParseCsvFile(NMDC_ExchangePath() & "\dashboard.csv", dashboardHeaders, dashboardData, dashboardRows, dashboardCols) Then Err.Raise vbObjectError + 223, , "Dashboard baseline is unavailable."
    If StrComp(CStr(NMDC_ReviewDraftCsvValue(dashboardHeaders, dashboardData, dashboardRows, "Pending Run ID")), runId, vbBinaryCompare) <> 0 Then
        ' A newer stage may already have replaced exchange files. Leave the old
        ' run's journal intact; never compare the old grid against a new baseline.
        NMDC_ReviewDraftCaptureCurrentTables = True
        Exit Function
    End If
    Set drafts = NMDC_ReviewDraftLoad()

    If Not NMDC_ParseCsvFile(NMDC_ExchangePath() & "\source_selection.csv", sourceHeaders, sourceData, sourceRows, sourceCols, True) Then Err.Raise vbObjectError + 226, , "Source selection baseline is unavailable."
    Set ws = ThisWorkbook.Worksheets("Pending Update")
        Set table = Nothing
        On Error Resume Next
        Set table = ws.ListObjects("SourceSelection")
        On Error GoTo Handler
        If Not table Is Nothing Then
            If Not table.DataBodyRange Is Nothing Then
                For rowIndex = 1 To table.ListRows.Count
                    sourceKey = NMDC_ReviewDraftNormalizeSource(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value2))
                    If Len(sourceKey) = 0 Then GoTo NextSourceBaselineRow
                    csvRow = NMDC_ReviewDraftFindSourceCsvRow(sourceData, sourceRows, sourceHeaders, sourceKey)
                    If csvRow = 0 Then Err.Raise vbObjectError + 229, , "Displayed source row is missing from the current exchange baseline."
                    If csvRow > 0 Then
                        For Each fieldName In Array("Include in Index?", "Owner Note")
                            If NMDC_ReviewDraftCsvHasField(sourceHeaders, CStr(fieldName)) Then
                                currentValue = CStr(table.ListColumns(CStr(fieldName)).DataBodyRange.Cells(rowIndex, 1).Value2)
                                baselineValue = NMDC_ReviewDraftCsvDataValue(sourceData, csvRow, sourceHeaders, CStr(fieldName))
                                If NMDC_ReviewDraftComparable(CStr(fieldName), currentValue) <> NMDC_ReviewDraftComparable(CStr(fieldName), baselineValue) Then
                                    NMDC_ReviewDraftStore drafts, runId, SOURCE_SECTION, sourceKey, CStr(fieldName), currentValue
                                    changed = True
                                Else
                                    If NMDC_ReviewDraftRemove(drafts, runId, SOURCE_SECTION, sourceKey, CStr(fieldName)) Then changed = True
                                End If
                            End If
                        Next fieldName
                    End If
NextSourceBaselineRow:
                Next rowIndex
            End If
        End If

    If Not NMDC_ParseCsvFile(NMDC_ExchangePath() & "\flags.csv", flagHeaders, flagData, flagRows, flagCols, True) Then Err.Raise vbObjectError + 227, , "Review Flags baseline is unavailable."
    Set ws = ThisWorkbook.Worksheets("Review Flags")
        Set table = ws.ListObjects("ReviewFlags")
        If Not table.DataBodyRange Is Nothing Then
            For rowIndex = 1 To table.ListRows.Count
                flagKey = NMDC_ReviewDraftFlagIdentity(table, rowIndex)
                If Len(NMDC_ReviewDraftCellText(table, rowIndex, "Flag Code")) = 0 Then GoTo NextFlagBaselineRow
                csvRow = NMDC_ReviewDraftFindFlagCsvRow(flagData, flagRows, flagHeaders, flagKey)
                If csvRow = 0 Then Err.Raise vbObjectError + 230, , "Displayed Review Flag is missing from the current exchange baseline."
                If csvRow > 0 Then
                    For Each fieldName In Array("Select?", "User Decision", "User Comment", "Resolution Status")
                        If NMDC_ReviewDraftCsvHasField(flagHeaders, CStr(fieldName)) Then
                            currentValue = CStr(table.ListColumns(CStr(fieldName)).DataBodyRange.Cells(rowIndex, 1).Value2)
                            baselineValue = NMDC_ReviewDraftCsvDataValue(flagData, csvRow, flagHeaders, CStr(fieldName))
                            If NMDC_ReviewDraftComparable(CStr(fieldName), currentValue) <> NMDC_ReviewDraftComparable(CStr(fieldName), baselineValue) Then
                                NMDC_ReviewDraftStore drafts, runId, FLAGS_SECTION, flagKey, CStr(fieldName), currentValue
                                changed = True
                            Else
                                If NMDC_ReviewDraftRemove(drafts, runId, FLAGS_SECTION, flagKey, CStr(fieldName)) Then changed = True
                            End If
                        End If
                    Next fieldName
                End If
NextFlagBaselineRow:
            Next rowIndex
        End If

    If changed Then NMDC_ReviewDraftWrite drafts
    mReviewDraftStoreFaulted = False
    NMDC_ReviewDraftCaptureCurrentTables = True
    Exit Function
Handler:
    mReviewDraftStoreFaulted = True
    NMDC_ReviewDraftReportStoreError "Excel could not compare unsaved review edits with the current proposal. Refresh and approval are blocked.", _
        Err.Number & " - " & Err.Description
    NMDC_ReviewDraftCaptureCurrentTables = False
End Function

' Called by the workbook SheetChange event. Only changed owner-input cells are journaled.
Public Sub NMDC_ReviewDraftSheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Handler

    Dim table As ListObject
    Dim body As Range
    Dim changed As Range
    Dim cell As Range
    Dim sectionText As String
    Dim fieldName As String
    Dim shouldCapture As Boolean
    Dim runId As String
    Dim baselineRunId As String
    Dim dashboardHeaders As Variant, dashboardData As Variant
    Dim dashboardRows As Long, dashboardCols As Long
    Dim drafts As Object
    Dim rowIndex As Long
    Dim identity As String
    Dim replacedBaseline As Boolean

    If Target Is Nothing Then Exit Sub
    Select Case CStr(Sh.Name)
        Case "Pending Update"
            Set table = Nothing
            On Error Resume Next
            Set table = Sh.ListObjects("SourceSelection")
            On Error GoTo Handler
            sectionText = SOURCE_SECTION
        Case "Review Flags"
            Set table = Nothing
            On Error Resume Next
            Set table = Sh.ListObjects("ReviewFlags")
            On Error GoTo Handler
            sectionText = FLAGS_SECTION
        Case Else
            Exit Sub
    End Select
    If table Is Nothing Then Exit Sub
    If table.DataBodyRange Is Nothing Then Exit Sub
    Set body = table.DataBodyRange
    Set changed = Intersect(Target, body)
    If changed Is Nothing Then Exit Sub
    runId = NMDC_DisplayedPendingRunId()
    If Len(runId) = 0 Then Exit Sub
    If Not NMDC_ParseCsvFile(NMDC_ExchangePath() & "\dashboard.csv", dashboardHeaders, dashboardData, dashboardRows, dashboardCols) Then Err.Raise vbObjectError + 231, , "Dashboard baseline is unavailable."
    baselineRunId = NMDC_ReviewDraftCsvValue(dashboardHeaders, dashboardData, dashboardRows, "Pending Run ID")
    replacedBaseline = (StrComp(runId, baselineRunId, vbBinaryCompare) <> 0)
    If replacedBaseline Then Set drafts = NMDC_ReviewDraftLoad()

    For Each cell In changed.Cells
        rowIndex = cell.Row - body.Row + 1
        fieldName = CStr(table.HeaderRowRange.Cells(1, cell.Column - table.Range.Column + 1).Value2)
        If sectionText = SOURCE_SECTION Then
            If fieldName <> "Include in Index?" And fieldName <> "Owner Note" Then GoTo NextCell
        Else
            If fieldName <> "Select?" And fieldName <> "User Decision" And _
               fieldName <> "User Comment" And fieldName <> "Resolution Status" Then GoTo NextCell
        End If
        If replacedBaseline Then
            If sectionText = SOURCE_SECTION Then
                identity = NMDC_ReviewDraftNormalizeSource(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value2))
            Else
                identity = NMDC_ReviewDraftFlagIdentity(table, rowIndex)
            End If
            If Len(identity) > 0 Then NMDC_ReviewDraftStore drafts, runId, sectionText, identity, fieldName, CStr(cell.Value2)
        Else
            shouldCapture = True
        End If
NextCell:
    Next cell
    If shouldCapture Then
        If Not NMDC_ReviewDraftCaptureCurrentTables() Then Err.Raise vbObjectError + 225, , "Draft comparison failed."
    End If
    If replacedBaseline Then NMDC_ReviewDraftWrite drafts
    mReviewDraftStoreFaulted = False
    Exit Sub

Handler:
    mReviewDraftStoreFaulted = True
    NMDC_ReviewDraftReportStoreError "Excel could not preserve an unsaved review edit. Stop and retry before refreshing or approving.", _
        Err.Number & " - " & Err.Description
End Sub

' Restore only draft cells belonging to the proposal currently displayed.
' Drafts from a replaced run remain in the CSV and are never applied to a new run.
Public Function NMDC_ReviewDraftsRestoreCurrent() As Boolean
    On Error GoTo Handler

    Dim drafts As Object
    Dim runId As String
    Dim key As Variant
    Dim item As Variant
    Dim ws As Worksheet
    Dim table As ListObject
    Dim rowIndex As Long
    Dim cell As Range
    Dim oldEvents As Boolean
    Dim dashboardHeaders As Variant, dashboardData As Variant
    Dim dashboardRows As Long, dashboardCols As Long
    Dim exchangeRunId As String
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    oldEvents = Application.EnableEvents
    Set drafts = NMDC_ReviewDraftLoad()
    runId = NMDC_DisplayedPendingRunId()
    If Len(runId) = 0 Then
        NMDC_ReviewDraftUpdateNotice drafts, ""
        NMDC_ReviewDraftAddReportButton
        NMDC_ReviewDraftsRestoreCurrent = True
        Exit Function
    End If
    If Not NMDC_ParseCsvFile(NMDC_ExchangePath() & "\dashboard.csv", dashboardHeaders, dashboardData, dashboardRows, dashboardCols) Then _
        Err.Raise vbObjectError + 241, , "The dashboard exchange file could not be parsed while restoring review drafts."
    If Not NMDC_ReviewDraftCsvHasField(dashboardHeaders, "Pending Run ID") Or dashboardRows < 1 Then _
        Err.Raise vbObjectError + 242, , "The dashboard exchange file is missing its pending-run baseline."
    exchangeRunId = NMDC_ReviewDraftCsvValue(dashboardHeaders, dashboardData, dashboardRows, "Pending Run ID")
    If StrComp(runId, exchangeRunId, vbBinaryCompare) <> 0 Then
        NMDC_ReviewDraftUpdateNotice drafts, exchangeRunId
        NMDC_ReviewDraftAddReportButton
        NMDC_ReviewDraftsRestoreCurrent = True
        Exit Function
    End If
    Application.EnableEvents = False

    For Each key In drafts.Keys
        item = drafts(key)
        If StrComp(CStr(item(0)), runId, vbBinaryCompare) = 0 Then
            Set ws = Nothing
            Set table = Nothing
            Select Case CStr(item(1))
                Case SOURCE_SECTION
                    Set ws = ThisWorkbook.Worksheets("Pending Update")
                    Set table = ws.ListObjects("SourceSelection")
                Case FLAGS_SECTION
                    Set ws = ThisWorkbook.Worksheets("Review Flags")
                    Set table = ws.ListObjects("ReviewFlags")
                Case Else
                    GoTo NextDraft
            End Select
            rowIndex = NMDC_ReviewDraftFindRow(table, CStr(item(2)), CStr(item(1)))
            If rowIndex > 0 Then
                Set cell = table.ListRows(rowIndex).Range.Cells(1, table.ListColumns(CStr(item(3))).Index)
                If CStr(item(3)) = "Include in Index?" Or CStr(item(3)) = "Select?" Then
                    cell.Value2 = NMDC_ReviewDraftBoolean(item(4))
                Else
                    cell.Value2 = CStr(item(4))
                End If
            End If
        End If
NextDraft:
    Next key

    Application.EnableEvents = oldEvents
    NMDC_ReviewDraftUpdateNotice drafts, runId
    NMDC_ReviewDraftAddReportButton
    NMDC_ReviewDraftsRestoreCurrent = True
    Exit Function

Handler:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    mReviewDraftStoreFaulted = True
    On Error Resume Next
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    NMDC_ReviewDraftReportStoreError "Excel could not restore the saved review draft. Refresh and approval are blocked until the draft store is readable.", _
        savedErrorNumber & " - " & savedErrorDescription
    NMDC_ReviewDraftsRestoreCurrent = False
End Function

Public Function NMDC_ReviewDraftsCanApprove(ByVal runId As String, ByRef reason As String) As Boolean
    On Error GoTo Handler

    Dim drafts As Object
    Dim key As Variant
    Dim item As Variant
    Dim sourceCount As Long
    Dim flagCount As Long
    Dim restagePending As Boolean

    Set drafts = NMDC_ReviewDraftLoad()
    For Each key In drafts.Keys
        item = drafts(key)
        If CStr(item(1)) = SOURCE_SECTION Then
            If StrComp(CStr(item(0)), runId, vbBinaryCompare) = 0 Then sourceCount = sourceCount + 1
        ElseIf CStr(item(1)) = STATE_SECTION And CStr(item(2)) = "SOURCE_RESTAGE" Then
            restagePending = True
        ElseIf CStr(item(1)) = FLAGS_SECTION And _
               StrComp(CStr(item(0)), runId, vbBinaryCompare) = 0 And _
               (CStr(item(3)) = "User Decision" Or CStr(item(3)) = "User Comment" Or CStr(item(3)) = "Resolution Status") Then
            flagCount = flagCount + 1
        End If
    Next key

    If sourceCount > 0 Or flagCount > 0 Or restagePending Then
        reason = "Approval is blocked because owner changes are still draft. "
        If sourceCount > 0 Then reason = reason & CStr(sourceCount) & " source choice/note field(s) must be saved and restaged. "
        If restagePending Then reason = reason & "Saved source choices await a successful new staged proposal. "
        If flagCount > 0 Then reason = reason & CStr(flagCount) & " Review Flag decision/comment/status field(s) must be saved. "
        reason = reason & "Use Review Saved Drafts to inspect preserved edits, then save the current review inputs."
        NMDC_ReviewDraftsCanApprove = False
        Exit Function
    End If
    NMDC_ReviewDraftsCanApprove = True
    Exit Function

Handler:
    reason = "Approval is blocked because Excel could not read the saved owner draft store. Review Error Log before continuing."
    NMDC_ReviewDraftsCanApprove = False
    NMDC_ReviewDraftReportStoreError reason, Err.Number & " - " & Err.Description
End Function

Public Function NMDC_ReviewDraftFlagIdentity(ByVal table As ListObject, ByVal rowIndex As Long) As String
    Dim fields As Variant
    Dim i As Long
    Dim valueText As String

    fields = Array("Flag Code", "Source File", "Worksheet Name", "Event Key", "Project No.", "Document No.", "Revision")
    For i = LBound(fields) To UBound(fields)
        valueText = NMDC_ReviewDraftCellText(table, rowIndex, CStr(fields(i)))
        If i = 0 Then
            NMDC_ReviewDraftFlagIdentity = NMDC_ReviewDraftPart(valueText)
        Else
            NMDC_ReviewDraftFlagIdentity = NMDC_ReviewDraftFlagIdentity & "|" & NMDC_ReviewDraftPart(valueText)
        End If
    Next i
End Function

Public Function NMDC_ReviewDraftGet(ByVal runId As String, ByVal sectionText As String, ByVal identity As String, _
                                    ByVal fieldName As String, ByRef found As Boolean) As String
    Dim drafts As Object
    Dim key As String
    Dim item As Variant
    Set drafts = NMDC_ReviewDraftLoad()
    key = NMDC_ReviewDraftKey(runId, sectionText, identity, fieldName)
    found = drafts.Exists(key)
    If found Then
        item = drafts(key)
        NMDC_ReviewDraftGet = CStr(item(4))
    Else
        NMDC_ReviewDraftGet = ""
    End If
End Function

Public Sub NMDC_ReviewDraftsClearFlagFields(ByVal runId As String)
    NMDC_ReviewDraftClearSection runId, FLAGS_SECTION, True
End Sub

Public Function NMDC_ReviewDraftIsDirty(ByVal runId As String, ByVal sectionText As String, _
                                        ByVal identity As String, ByVal fieldName As String, _
                                        ByRef draftValue As String) As Boolean
    Dim found As Boolean
    draftValue = NMDC_ReviewDraftGet(runId, sectionText, identity, fieldName, found)
    NMDC_ReviewDraftIsDirty = found
End Function

Public Function NMDC_ReviewDraftsSnapshot() As Object
    Set NMDC_ReviewDraftsSnapshot = NMDC_ReviewDraftLoad()
End Function

Public Function NMDC_ReviewDraftSnapshotIsDirty(ByVal drafts As Object, ByVal runId As String, _
                                                  ByVal sectionText As String, ByVal identity As String, _
                                                  ByVal fieldName As String, ByRef draftValue As String) As Boolean
    Dim key As String
    Dim item As Variant
    key = NMDC_ReviewDraftKey(runId, sectionText, identity, fieldName)
    NMDC_ReviewDraftSnapshotIsDirty = drafts.Exists(key)
    If NMDC_ReviewDraftSnapshotIsDirty Then
        item = drafts(key)
        draftValue = CStr(item(4))
    Else
        draftValue = ""
    End If
End Function

Public Sub NMDC_ReviewDraftCaptureAllSourceChoices()
    If Not NMDC_ReviewDraftCaptureCurrentTables() Then
        NMDC_ReviewDraftReportStoreError "Excel could not preserve the bulk source choices. Stop before refreshing or approving.", _
            "The displayed owner inputs could not be compared with the current proposal."
    End If
End Sub

Public Sub NMDC_ReviewDraftMarkSourceRestage(ByVal runId As String)
    NMDC_ReviewDraftSet runId, STATE_SECTION, "SOURCE_RESTAGE", "WAITING", "TRUE"
End Sub

Public Sub NMDC_ReviewDraftCompleteSourceRestage()
    On Error GoTo Handler
    Dim drafts As Object
    Dim keys As Variant
    Dim key As Variant
    Dim item As Variant
    Dim currentRunId As String
    Dim dashboardHeaders As Variant, dashboardData As Variant
    Dim sourceHeaders As Variant, sourceData As Variant
    Dim dashboardRows As Long, dashboardCols As Long
    Dim sourceRows As Long, sourceCols As Long
    Dim sourceKey As String
    Dim csvRow As Long
    Dim fieldName As Variant
    Dim changed As Boolean
    Dim oldRunId As String

    Set drafts = NMDC_ReviewDraftLoad()
    currentRunId = NMDC_DisplayedPendingRunId()
    If Len(currentRunId) = 0 Then Exit Sub
    If Not NMDC_ParseCsvFile(NMDC_ExchangePath() & "\dashboard.csv", dashboardHeaders, dashboardData, dashboardRows, dashboardCols) Then _
        Err.Raise vbObjectError + 236, , "The new staged proposal dashboard is unavailable."
    If StrComp(NMDC_ReviewDraftCsvValue(dashboardHeaders, dashboardData, dashboardRows, "Pending Run ID"), _
               currentRunId, vbBinaryCompare) <> 0 Then Exit Sub
    If Not NMDC_ParseCsvFile(NMDC_ExchangePath() & "\source_selection.csv", sourceHeaders, sourceData, sourceRows, sourceCols, True) Then _
        Err.Raise vbObjectError + 237, , "The new staged source baseline is unavailable."
    If Not NMDC_ReviewDraftCsvHasField(sourceHeaders, "Include in Index?") Or _
       Not NMDC_ReviewDraftCsvHasField(sourceHeaders, "Owner Note") Then _
        Err.Raise vbObjectError + 238, , "The new staged source baseline is missing owner-input fields."

    keys = drafts.Keys
    For Each key In keys
        If drafts.Exists(CStr(key)) Then
            item = drafts(key)
            If CStr(item(1)) = STATE_SECTION And CStr(item(2)) = "SOURCE_RESTAGE" And _
               Len(currentRunId) > 0 And StrComp(currentRunId, CStr(item(0)), vbBinaryCompare) <> 0 Then
                oldRunId = CStr(item(0))
                For Each fieldName In keys
                    If drafts.Exists(CStr(fieldName)) Then
                        item = drafts(fieldName)
                        If CStr(item(0)) = oldRunId And CStr(item(1)) = SOURCE_SECTION Then
                            sourceKey = CStr(item(2))
                            csvRow = NMDC_ReviewDraftFindSourceCsvRow(sourceData, sourceRows, sourceHeaders, sourceKey)
                            If csvRow > 0 Then
                                If NMDC_ReviewDraftComparable(CStr(item(3)), CStr(item(4))) = _
                                   NMDC_ReviewDraftComparable(CStr(item(3)), NMDC_ReviewDraftCsvDataValue(sourceData, csvRow, sourceHeaders, CStr(item(3)))) Then
                                    drafts.Remove CStr(fieldName)
                                    changed = True
                                End If
                            End If
                        End If
                    End If
                Next fieldName

                ' Keep the restage marker while any old-run source edit is unmatched.
                ' This protects edits made after Save was clicked or when the saved
                ' configuration did not make it into the staged proposal.
                If Not NMDC_ReviewDraftHasSourceFields(drafts, oldRunId) Then
                    If NMDC_ReviewDraftRemove(drafts, oldRunId, STATE_SECTION, "SOURCE_RESTAGE", "WAITING") Then changed = True
                End If
            End If
        End If
    Next key
    If changed Then NMDC_ReviewDraftWrite drafts
    NMDC_ReviewDraftUpdateNotice drafts, currentRunId
    NMDC_ReviewDraftAddReportButton
    Exit Sub
Handler:
    mReviewDraftStoreFaulted = True
    NMDC_ReviewDraftReportStoreError "Excel could not clear source drafts after a successful restage.", Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_ReviewDraftHasSourceFields(ByVal drafts As Object, ByVal runId As String) As Boolean
    Dim key As Variant
    Dim item As Variant
    For Each key In drafts.Keys
        item = drafts(key)
        If CStr(item(0)) = runId And CStr(item(1)) = SOURCE_SECTION Then
            NMDC_ReviewDraftHasSourceFields = True
            Exit Function
        End If
    Next key
End Function

Public Sub NMDC_ReviewDraftOpenReport()
    On Error GoTo Handler
    Dim path As String
    path = NMDC_ReviewDraftPath()
    If Not NMDC_FileExists(path) Then
        MsgBox "There are no saved owner drafts to review.", vbInformation, "Review Saved Drafts"
        Exit Sub
    End If
    Application.Workbooks.Open Filename:=path, ReadOnly:=True
    Exit Sub
Handler:
    NMDC_ReviewDraftReportStoreError "Excel could not open the saved owner draft report.", Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_ReviewDraftSet(ByVal runId As String, ByVal sectionText As String, ByVal identity As String, _
                               ByVal fieldName As String, ByVal valueText As String)
    Dim drafts As Object
    Set drafts = NMDC_ReviewDraftLoad()
    NMDC_ReviewDraftStore drafts, runId, sectionText, identity, fieldName, valueText
    NMDC_ReviewDraftWrite drafts
End Sub

Private Sub NMDC_ReviewDraftStore(ByVal drafts As Object, ByVal runId As String, ByVal sectionText As String, _
                                  ByVal identity As String, ByVal fieldName As String, ByVal valueText As String)
    drafts(NMDC_ReviewDraftKey(runId, sectionText, identity, fieldName)) = _
        Array(runId, sectionText, identity, fieldName, valueText)
End Sub

Private Function NMDC_ReviewDraftLoad() As Object
    On Error GoTo Handler
    Dim drafts As Object
    Dim headers As Variant
    Dim data As Variant
    Dim count As Long
    Dim columns As Long
    Dim i As Long
    Dim rowItem As Variant
    Dim filePath As String
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String
    Set drafts = CreateObject("Scripting.Dictionary")
    drafts.CompareMode = vbBinaryCompare
    filePath = NMDC_ReviewDraftPath()
    If NMDC_FileExists(filePath & ".tmp") Or NMDC_FileExists(filePath & ".bak") Then _
        Err.Raise vbObjectError + 239, , "The owner draft store has an interrupted write (.tmp/.bak present); preserve these files and review before continuing."
    If NMDC_FileExists(filePath) Then
        If Not NMDC_ParseCsvFile(filePath, headers, data, count, columns, True) Then Err.Raise vbObjectError + 221, , "The owner draft CSV could not be parsed."
        If columns <> 5 Or CStr(headers(1, 1)) <> "Run ID" Or CStr(headers(1, 2)) <> "Section" Or _
           CStr(headers(1, 3)) <> "Identity" Or CStr(headers(1, 4)) <> "Field" Or CStr(headers(1, 5)) <> "Value" Then
            Err.Raise vbObjectError + 222, , "The owner draft CSV has an unsupported header."
        End If
        For i = 1 To count
            rowItem = Array(CStr(data(i, 1)), CStr(data(i, 2)), CStr(data(i, 3)), CStr(data(i, 4)), CStr(data(i, 5)))
            If Not NMDC_ReviewDraftRowValid(rowItem) Then Err.Raise vbObjectError + 233, , "The owner draft CSV contains an unsupported or incomplete row."
            If drafts.Exists(NMDC_ReviewDraftKey(rowItem(0), rowItem(1), rowItem(2), rowItem(3))) Then _
                Err.Raise vbObjectError + 234, , "The owner draft CSV contains a duplicate row identity."
            drafts.Add NMDC_ReviewDraftKey(rowItem(0), rowItem(1), rowItem(2), rowItem(3)), rowItem
        Next i
    End If
    Set NMDC_ReviewDraftLoad = drafts
    Exit Function
Handler:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    mReviewDraftStoreFaulted = True
    Err.Raise savedErrorNumber, "NMDC_ReviewDraftLoad", savedErrorDescription
End Function

Private Sub NMDC_ReviewDraftWrite(ByVal drafts As Object)
    On Error GoTo Handler
    Dim stream As Object
    Dim fso As Object
    Dim key As Variant
    Dim item As Variant
    Dim path As String
    Dim tempPath As String
    Dim backupPath As String
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    path = NMDC_ReviewDraftPath()
    tempPath = path & ".tmp"
    backupPath = path & ".bak"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(NMDC_RuntimePath()) Then fso.CreateFolder NMDC_RuntimePath()
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText DRAFT_HEADERS & vbLf
    For Each key In drafts.Keys
        item = drafts(key)
        If Not NMDC_ReviewDraftRowValid(item) Then Err.Raise vbObjectError + 235, , "An owner draft row is incomplete or unsupported."
        stream.WriteText NMDC_ReviewDraftCsvField(item(0)) & "," & NMDC_ReviewDraftCsvField(item(1)) & "," & _
                          NMDC_ReviewDraftCsvField(item(2)) & "," & NMDC_ReviewDraftCsvField(item(3)) & "," & _
                          NMDC_ReviewDraftCsvField(item(4)) & vbLf
    Next key
    stream.SaveToFile tempPath, 2
    stream.Close

    If fso.FileExists(backupPath) Then fso.DeleteFile backupPath, True
    If fso.FileExists(path) Then fso.MoveFile path, backupPath
    fso.MoveFile tempPath, path
    If fso.FileExists(backupPath) Then fso.DeleteFile backupPath, True
    Exit Sub
Handler:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    mReviewDraftStoreFaulted = True
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Not fso Is Nothing Then
        If fso.FileExists(backupPath) And Not fso.FileExists(path) Then fso.MoveFile backupPath, path
    End If
    On Error GoTo 0
    Err.Raise savedErrorNumber, "NMDC_ReviewDraftWrite", savedErrorDescription
End Sub

Private Function NMDC_ReviewDraftRowValid(ByVal item As Variant) As Boolean
    Dim sectionText As String
    Dim fieldName As String
    If Len(Trim$(CStr(item(0)))) = 0 Or Len(Trim$(CStr(item(2)))) = 0 Then Exit Function
    sectionText = CStr(item(1))
    fieldName = CStr(item(3))
    Select Case sectionText
        Case SOURCE_SECTION
            If fieldName = "Include in Index?" Then
                NMDC_ReviewDraftRowValid = NMDC_ReviewDraftBooleanValueValid(CStr(item(4)))
            Else
                NMDC_ReviewDraftRowValid = (fieldName = "Owner Note")
            End If
        Case FLAGS_SECTION
            If fieldName = "Select?" Then
                NMDC_ReviewDraftRowValid = NMDC_ReviewDraftBooleanValueValid(CStr(item(4)))
            Else
                NMDC_ReviewDraftRowValid = (fieldName = "User Decision" Or _
                    fieldName = "User Comment" Or fieldName = "Resolution Status")
            End If
        Case STATE_SECTION
            NMDC_ReviewDraftRowValid = (CStr(item(2)) = "SOURCE_RESTAGE" And _
                fieldName = "WAITING" And UCase$(Trim$(CStr(item(4)))) = "TRUE")
    End Select
End Function

Private Function NMDC_ReviewDraftBooleanValueValid(ByVal valueText As String) As Boolean
    Select Case UCase$(Trim$(valueText))
        Case "TRUE", "FALSE", "YES", "NO", "1", "0"
            NMDC_ReviewDraftBooleanValueValid = True
    End Select
End Function

Private Sub NMDC_ReviewDraftClearSection(ByVal runId As String, ByVal sectionText As String, ByVal decisionsOnly As Boolean)
    Dim drafts As Object
    Dim keys As Variant
    Dim i As Long
    Dim item As Variant
    Set drafts = NMDC_ReviewDraftLoad()
    If drafts.Count = 0 Then Exit Sub
    keys = drafts.Keys
    For i = LBound(keys) To UBound(keys)
        item = drafts(keys(i))
        If CStr(item(1)) = sectionText And (Len(runId) = 0 Or CStr(item(0)) = runId) Then
            If Not decisionsOnly Or CStr(item(3)) = "User Decision" Or _
               CStr(item(3)) = "User Comment" Or CStr(item(3)) = "Resolution Status" Then
                drafts.Remove CStr(keys(i))
            End If
        End If
    Next i
    NMDC_ReviewDraftWrite drafts
End Sub

Private Function NMDC_ReviewDraftFindRow(ByVal table As ListObject, ByVal identity As String, ByVal sectionText As String) As Long
    Dim rowIndex As Long
    Dim candidate As String
    If table.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = 1 To table.ListRows.Count
        If sectionText = SOURCE_SECTION Then
            candidate = NMDC_ReviewDraftNormalizeSource(CStr(table.ListColumns("Source File").DataBodyRange.Cells(rowIndex, 1).Value2))
        Else
            candidate = NMDC_ReviewDraftFlagIdentity(table, rowIndex)
        End If
        If StrComp(candidate, identity, vbBinaryCompare) = 0 Then
            NMDC_ReviewDraftFindRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub NMDC_ReviewDraftUpdateNotice(ByVal drafts As Object, ByVal currentRunId As String)
    On Error Resume Next
    Dim key As Variant
    Dim item As Variant
    Dim staleRuns As Object
    Dim staleCount As Long
    Dim currentCount As Long
    Dim runKey As Variant
    Dim notice As String
    Set staleRuns = CreateObject("Scripting.Dictionary")
    For Each key In drafts.Keys
        item = drafts(key)
        If CStr(item(1)) <> STATE_SECTION Then
            If StrComp(CStr(item(0)), currentRunId, vbBinaryCompare) <> 0 Then
                staleCount = staleCount + 1
                If Not staleRuns.Exists(CStr(item(0))) Then staleRuns.Add CStr(item(0)), True
            Else
                currentCount = currentCount + 1
            End If
        End If
    Next key
    If currentCount > 0 Then
        notice = " UNSAVED DRAFTS: " & CStr(currentCount) & " field edit(s) are attached to this proposal and must be saved before approval."
    End If
    If staleCount > 0 Then
        notice = notice & " SAVED DRAFT WARNING: " & CStr(staleCount) & " edit(s) from replaced proposal run(s) "
        For Each runKey In staleRuns.Keys
            notice = notice & CStr(runKey) & " "
        Next runKey
        notice = notice & "were retained, not applied. Use Review Saved Drafts to inspect before continuing."
    End If
    With ThisWorkbook.Worksheets("Pending Update").Range("A4")
        If currentCount > 0 Or staleCount > 0 Then
            .Value = "REVIEW WORKFLOW. Source choices and notes are saved by Save Source Choices & Restage; flag decisions/comments/status are saved by Save Review Decisions. Approve/Hold/Reject applies to the whole proposal." & notice
            .WrapText = True
        Else
            .Value = "REVIEW WORKFLOW. Review staged changes first. Source checkboxes and Owner Notes are drafts until Save Source Choices & Restage. Approve/Hold/Reject applies to the whole proposal."
        End If
    End With
End Sub

Private Sub NMDC_ReviewDraftAddReportButton()
    On Error Resume Next
    Dim ws As Worksheet
    Dim table As ListObject
    Dim area As Range
    Dim button As Shape
    Set ws = ThisWorkbook.Worksheets("Pending Update")
    Set table = ws.ListObjects("SourceSelection")
    If table Is Nothing Then Exit Sub
    ws.Shapes("NMDC_Pending_ReviewDrafts").Delete
    Set area = ws.Range(ws.Cells(table.HeaderRowRange.Row - 1, 9), ws.Cells(table.HeaderRowRange.Row - 1, 10))
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Pending_ReviewDrafts"
    button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!NMDC_ReviewDraftOpenReport"
    button.TextFrame.Characters.Text = "Review Saved Drafts"
    button.TextFrame.HorizontalAlignment = xlCenter
    button.TextFrame.VerticalAlignment = 3
    button.Fill.ForeColor.RGB = RGB(91, 100, 112)
    button.Line.ForeColor.RGB = RGB(91, 100, 112)
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 8.5
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    button.Placement = xlMoveAndSize
End Sub

Private Function NMDC_ReviewDraftPath() As String
    NMDC_ReviewDraftPath = NMDC_RuntimePath() & "\" & DRAFT_FILE
End Function

Private Function NMDC_ReviewDraftCsvValue(ByVal headers As Variant, ByVal data As Variant, ByVal rowCount As Long, ByVal fieldName As String) As String
    Dim columnIndex As Long
    columnIndex = NMDC_ReviewDraftCsvColumn(headers, fieldName)
    If columnIndex > 0 And rowCount > 0 Then NMDC_ReviewDraftCsvValue = CStr(data(1, columnIndex))
End Function

Private Function NMDC_ReviewDraftCsvHasField(ByVal headers As Variant, ByVal fieldName As String) As Boolean
    NMDC_ReviewDraftCsvHasField = (NMDC_ReviewDraftCsvColumn(headers, fieldName) > 0)
End Function

Private Function NMDC_ReviewDraftCsvColumn(ByVal headers As Variant, ByVal fieldName As String) As Long
    Dim i As Long
    For i = 1 To UBound(headers, 2)
        If StrComp(Trim$(CStr(headers(1, i))), fieldName, vbTextCompare) = 0 Then
            NMDC_ReviewDraftCsvColumn = i
            Exit Function
        End If
    Next i
End Function

Private Function NMDC_ReviewDraftCsvDataValue(ByVal data As Variant, ByVal rowIndex As Long, ByVal headers As Variant, ByVal fieldName As String) As String
    Dim columnIndex As Long
    columnIndex = NMDC_ReviewDraftCsvColumn(headers, fieldName)
    If columnIndex > 0 Then NMDC_ReviewDraftCsvDataValue = CStr(data(rowIndex, columnIndex))
End Function

Private Function NMDC_ReviewDraftFindSourceCsvRow(ByVal data As Variant, ByVal rowCount As Long, ByVal headers As Variant, ByVal sourceKey As String) As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    columnIndex = NMDC_ReviewDraftCsvColumn(headers, "Source File")
    If columnIndex = 0 Then columnIndex = NMDC_ReviewDraftCsvColumn(headers, "Relative Path")
    If columnIndex = 0 Then Err.Raise vbObjectError + 228, , "Source selection baseline has no source path column."
    For rowIndex = 1 To rowCount
        If NMDC_ReviewDraftNormalizeSource(CStr(data(rowIndex, columnIndex))) = sourceKey Then
            NMDC_ReviewDraftFindSourceCsvRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function NMDC_ReviewDraftFindFlagCsvRow(ByVal data As Variant, ByVal rowCount As Long, ByVal headers As Variant, ByVal flagKey As String) As Long
    Dim rowIndex As Long
    Dim fields As Variant
    Dim i As Long
    Dim candidate As String
    Dim valueText As String
    fields = Array("Flag Code", "Source File", "Worksheet Name", "Event Key", "Project No.", "Document No.", "Revision")
    For rowIndex = 1 To rowCount
        candidate = ""
        For i = LBound(fields) To UBound(fields)
            valueText = NMDC_ReviewDraftCsvDataValue(data, rowIndex, headers, CStr(fields(i)))
            If i = LBound(fields) Then
                candidate = NMDC_ReviewDraftPart(valueText)
            Else
                candidate = candidate & "|" & NMDC_ReviewDraftPart(valueText)
            End If
        Next i
        If candidate = flagKey Then
            NMDC_ReviewDraftFindFlagCsvRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function NMDC_ReviewDraftComparable(ByVal fieldName As String, ByVal valueText As String) As String
    If fieldName = "Include in Index?" Or fieldName = "Select?" Then
        If NMDC_ReviewDraftBoolean(valueText) Then
            NMDC_ReviewDraftComparable = "TRUE"
        Else
            NMDC_ReviewDraftComparable = "FALSE"
        End If
    ElseIf fieldName = "User Decision" Or fieldName = "Resolution Status" Then
        NMDC_ReviewDraftComparable = UCase$(Trim$(valueText))
    Else
        NMDC_ReviewDraftComparable = valueText
    End If
End Function

Private Function NMDC_ReviewDraftRemove(ByVal drafts As Object, ByVal runId As String, ByVal sectionText As String, ByVal identity As String, ByVal fieldName As String) As Boolean
    Dim key As String
    key = NMDC_ReviewDraftKey(runId, sectionText, identity, fieldName)
    If drafts.Exists(key) Then
        drafts.Remove key
        NMDC_ReviewDraftRemove = True
    End If
End Function

Private Function NMDC_ReviewDraftKey(ByVal runId As String, ByVal sectionText As String, ByVal identity As String, ByVal fieldName As String) As String
    NMDC_ReviewDraftKey = runId & ChrW$(30) & sectionText & ChrW$(30) & identity & ChrW$(30) & fieldName
End Function

Private Function NMDC_ReviewDraftCellText(ByVal table As ListObject, ByVal rowIndex As Long, ByVal fieldName As String) As String
    On Error GoTo Missing
    NMDC_ReviewDraftCellText = CStr(table.ListColumns(fieldName).DataBodyRange.Cells(rowIndex, 1).Value2)
    Exit Function
Missing:
    NMDC_ReviewDraftCellText = ""
End Function

Private Function NMDC_ReviewDraftPart(ByVal valueText As String) As String
    NMDC_ReviewDraftPart = CStr(Len(valueText)) & ":" & valueText
End Function

Private Function NMDC_ReviewDraftNormalizeSource(ByVal valueText As String) As String
    NMDC_ReviewDraftNormalizeSource = LCase$(Replace(Trim$(valueText), "\", "/"))
End Function

Private Function NMDC_ReviewDraftBoolean(ByVal value As Variant) As Boolean
    NMDC_ReviewDraftBoolean = (UCase$(Trim$(CStr(value))) = "TRUE" Or UCase$(Trim$(CStr(value))) = "YES" Or CStr(value) = "1")
End Function

Private Function NMDC_ReviewDraftCsvField(ByVal value As Variant) As String
    Dim text As String
    text = CStr(value)
    If InStr(text, Chr$(34)) > 0 Then text = Replace(text, Chr$(34), Chr$(34) & Chr$(34))
    If InStr(text, ",") > 0 Or InStr(text, Chr$(34)) > 0 Or InStr(text, vbLf) > 0 Or InStr(text, vbCr) > 0 Then
        text = Chr$(34) & text & Chr$(34)
    End If
    NMDC_ReviewDraftCsvField = text
End Function

Private Sub NMDC_ReviewDraftReportStoreError(ByVal message As String, ByVal detail As String)
    On Error Resume Next
    NMDC_LogError "OWNER_REVIEW_DRAFT_ERROR", message, detail
    MsgBox message & vbCrLf & vbCrLf & detail, vbCritical, "NMDC Document Index"
    On Error GoTo 0
End Sub
