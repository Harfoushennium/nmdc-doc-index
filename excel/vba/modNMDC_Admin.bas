Attribute VB_Name = "modNMDC_Admin"
Option Explicit

Public Sub NMDC_ResetAllRecords()
    Dim answer As VbMsgBoxResult
    Dim exitCode As Long

    answer = MsgBox("Reset the NMDC Document Index and delete all indexed/staged runtime records?" & vbCrLf & vbCrLf & _
                    "This does NOT delete or modify source DATA workbooks or configuration files." & vbCrLf & _
                    "Audit history is preserved. You will need to run Full Rescan to rebuild the index.", _
                    vbCritical + vbYesNo + vbDefaultButton2, "Reset All Records")
    If answer <> vbYes Then Exit Sub

    answer = MsgBox("Confirm RESET ALL RECORDS." & vbCrLf & vbCrLf & _
                    "Approved records, pending updates, extraction cache and user flags will be cleared.", _
                    vbCritical + vbYesNo + vbDefaultButton2, "Final Reset Confirmation")
    If answer <> vbYes Then Exit Sub

    exitCode = NMDC_RunEngine("reset")
    If exitCode <> 0 Then
        MsgBox "Reset could not be completed. Please review the Error Log.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    If NMDC_RunEngine("export-excel") = 0 Then NMDC_RefreshExchangeData
    MsgBox "All indexed/staged records were reset." & vbCrLf & _
           "Source DATA and configuration were not changed. Run Full Rescan / Rebuild All when ready.", _
           vbInformation, "NMDC Document Index"
End Sub

Public Sub NMDC_UndoLastApproval()
    Dim answer As VbMsgBoxResult
    Dim exitCode As Long

    answer = MsgBox("Undo the last approved update and restore the previous approved version?" & vbCrLf & vbCrLf & _
                    "No source workbook will be changed or deleted. Previous approved versions are retained for audit.", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "Undo Last Approval")
    If answer <> vbYes Then Exit Sub

    exitCode = NMDC_RunEngine("undo")
    If exitCode <> 0 Then
        MsgBox "There is no previous approved version to restore, or Undo could not be completed." & vbCrLf & _
               "Please review the Error Log for details.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    If NMDC_RunEngine("export-excel") = 0 Then NMDC_RefreshExchangeData
    MsgBox "The last approval was undone and the previous approved version was restored.", _
           vbInformation, "NMDC Document Index"
End Sub

Public Sub NMDC_SaveReviewDecisions()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim row As ListRow
    Dim stream As Object
    Dim filePath As String
    Dim decision As String
    Dim comment As String
    Dim status As String
    Dim savedRows As Long
    Dim fixRows As Long
    Dim exitCode As Long
    Dim requestPath As String

    Set ws = ThisWorkbook.Worksheets("Review Flags")
    Set table = ws.ListObjects("ReviewFlags")
    filePath = NMDC_ExchangePath() & "\review_decisions.csv"
    NMDC_EnsureAdminFolder NMDC_ExchangePath()

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText "Flag Code,Source File,Worksheet Name,Event Key,Project No.,Document No.,Revision,User Decision,User Comment,Resolution Status" & vbLf

    For Each row In table.ListRows
        decision = Trim$(CStr(NMDC_AdminTableValue(table, row, "User Decision")))
        comment = Trim$(CStr(NMDC_AdminTableValue(table, row, "User Comment")))
        status = UCase$(Trim$(CStr(NMDC_AdminTableValue(table, row, "Resolution Status"))))

        If UCase$(decision) = "NEEDS PARSER/MAPPING FIX" Then fixRows = fixRows + 1

        If Len(decision) > 0 Or Len(comment) > 0 Or (Len(status) > 0 And status <> "OPEN") Then
            stream.WriteText _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Flag Code"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Source File"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Worksheet Name"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Event Key"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Project No."))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Document No."))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Revision"))) & "," & _
                NMDC_AdminCsvField(decision) & "," & _
                NMDC_AdminCsvField(comment) & "," & _
                NMDC_AdminCsvField(status) & vbLf
            savedRows = savedRows + 1
        End If
    Next row

    stream.SaveToFile filePath, 2
    stream.Close

    If savedRows = 0 Then
        MsgBox "No Review Flag decision/comment changes were found to save." & vbCrLf & vbCrLf & _
               "Use the User Decision dropdown and/or enter a User Comment first.", _
               vbInformation, "NMDC Document Index"
        Exit Sub
    End If

    exitCode = NMDC_RunEngine("save-review-decisions", _
        "--decisions-file " & NMDC_Quote(filePath) & _
        " --request-dir " & NMDC_Quote(NMDC_WorkbookFolder()))
    If exitCode <> 0 Then
        MsgBox "Review decisions could not be saved. Please review the Error Log.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    If NMDC_RunEngine("export-excel") = 0 Then NMDC_RefreshExchangeData

    If fixRows > 0 Then
        requestPath = NMDC_LatestParserFixRequestPath()
        MsgBox CStr(savedRows) & " Review Flag decision(s) were saved." & vbCrLf & vbCrLf & _
               CStr(fixRows) & " parser/mapping fix request(s) were prepared in a dedicated numbered folder:" & vbCrLf & _
               requestPath & vbCrLf & vbCrLf & _
               "Upload this report together with the affected source workbook(s) to ChatGPT / the project maintainer. " & _
               "After the corrected parser or mapping is installed, click Retry After Fix.", _
               vbInformation, "NMDC Document Index"
        NMDC_RevealFile requestPath
    Else
        MsgBox CStr(savedRows) & " Review Flag decision(s) were saved to the staged update audit trail.", _
               vbInformation, "NMDC Document Index"
    End If
    NMDC_GoToSheet "Review Flags"
    Exit Sub

Handler:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    NMDC_LogError "REVIEW_DECISION_SAVE_ERROR", _
        "Excel could not save the Review Flags decisions.", _
        Err.Number & " - " & Err.Description
    MsgBox "Review decisions could not be saved. Please review the Error Log.", vbExclamation, "NMDC Document Index"
End Sub

Public Sub NMDC_ReportSelectedParserFixes()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim row As ListRow
    Dim stream As Object
    Dim filePath As String
    Dim requestPath As String
    Dim selectedCount As Long
    Dim exitCode As Long
    Dim userNote As String
    Dim rowComment As String

    Set ws = ThisWorkbook.Worksheets("Review Flags")
    Set table = ws.ListObjects("ReviewFlags")

    If table.DataBodyRange Is Nothing Then
        MsgBox "There are no Review Flags to report.", vbInformation, "NMDC Document Index"
        Exit Sub
    End If

    For Each row In table.ListRows
        If NMDC_AdminCheckedValue(NMDC_AdminTableValue(table, row, "Select?")) Then
            selectedCount = selectedCount + 1
        End If
    Next row

    If selectedCount = 0 Then
        MsgBox "Tick the Select? checkbox for one or more Review Flag rows first." & vbCrLf & vbCrLf & _
               "Then click Report Selected Parser Fix.", _
               vbInformation, "NMDC Document Index"
        Exit Sub
    End If

    userNote = InputBox( _
        "Describe what the parser should extract for the selected flagged row(s)." & vbCrLf & vbCrLf & _
        "This note fills only blank User Comment cells. Existing row comments are kept." & vbCrLf & vbCrLf & _
        "Example: This is a normal document register. Document number is in column C and data begins at row 7.", _
        "Report Selected Parser Fix")
    If Len(Trim$(userNote)) = 0 Then
        userNote = "Please inspect this flagged source/sheet and update the parser or mapping so its normal register data is extracted."
    End If

    filePath = NMDC_ExchangePath() & "\review_decisions_selected.csv"
    NMDC_EnsureAdminFolder NMDC_ExchangePath()

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText "Flag Code,Source File,Worksheet Name,Event Key,Project No.,Document No.,Revision,User Decision,User Comment,Resolution Status" & vbLf

    For Each row In table.ListRows
        If NMDC_AdminCheckedValue(NMDC_AdminTableValue(table, row, "Select?")) Then
            row.Range.Cells(1, table.ListColumns("User Decision").Index).Value = "NEEDS PARSER/MAPPING FIX"
            row.Range.Cells(1, table.ListColumns("Resolution Status").Index).Value = "OPEN"

            rowComment = Trim$(CStr(NMDC_AdminTableValue(table, row, "User Comment")))
            If Len(rowComment) = 0 Then
                rowComment = userNote
                row.Range.Cells(1, table.ListColumns("User Comment").Index).Value = rowComment
            End If

            stream.WriteText _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Flag Code"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Source File"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Worksheet Name"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Event Key"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Project No."))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Document No."))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Revision"))) & "," & _
                NMDC_AdminCsvField("NEEDS PARSER/MAPPING FIX") & "," & _
                NMDC_AdminCsvField(rowComment) & "," & _
                NMDC_AdminCsvField("OPEN") & vbLf
        End If
    Next row

    stream.SaveToFile filePath, 2
    stream.Close

    exitCode = NMDC_RunEngine("save-review-decisions", _
        "--decisions-file " & NMDC_Quote(filePath) & _
        " --request-dir " & NMDC_Quote(NMDC_WorkbookFolder()))
    If exitCode <> 0 Then
        MsgBox "The selected parser-fix report could not be created. Please review Error Log.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    requestPath = NMDC_LatestParserFixRequestPath()

    If NMDC_RunEngine("export-excel") = 0 Then NMDC_RefreshExchangeData

    MsgBox CStr(selectedCount) & " selected parser/mapping issue(s) were reported." & vbCrLf & vbCrLf & _
           "The report is saved in its own numbered folder under PARSER_FIX_REPORTS:" & vbCrLf & _
           requestPath & vbCrLf & vbCrLf & _
           "Next: upload this report and the affected source workbook(s) to ChatGPT / the project maintainer. " & _
           "After a corrected parser/configuration is installed, click Retry After Fix.", _
           vbInformation, "NMDC Document Index"
    NMDC_RevealFile requestPath
    NMDC_GoToSheet "Review Flags"
    Exit Sub

Handler:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    NMDC_LogError "PARSER_FIX_REQUEST_ERROR", _
        "Excel could not prepare the selected parser/mapping fix report.", _
        Err.Number & " - " & Err.Description
    MsgBox "The selected parser-fix report could not be prepared. Please review Error Log.", _
           vbExclamation, "NMDC Document Index"
End Sub

Public Sub NMDC_RequestParserMappingFix()
    ' Backward-compatible action: the owner now selects one or more flags by checkbox.
    NMDC_ReportSelectedParserFixes
End Sub

Public Sub NMDC_SelectAllReviewFlags()
    NMDC_SetAllReviewFlagSelections True
End Sub

Public Sub NMDC_ClearReviewFlagSelection()
    NMDC_SetAllReviewFlagSelections False
End Sub

Private Sub NMDC_SetAllReviewFlagSelections(ByVal checkedValue As Boolean)
    On Error GoTo Handler

    Dim table As ListObject
    Dim row As ListRow
    Set table = ThisWorkbook.Worksheets("Review Flags").ListObjects("ReviewFlags")
    If table.DataBodyRange Is Nothing Then Exit Sub

    For Each row In table.ListRows
        If Len(Trim$(CStr(NMDC_AdminTableValue(table, row, "Flag Code")))) > 0 Then
            row.Range.Cells(1, table.ListColumns("Select?").Index).Value = checkedValue
        End If
    Next row
    Exit Sub

Handler:
    NMDC_LogError "REVIEW_SELECTION_ERROR", _
        "Excel could not change the Review Flags selection checkboxes.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_RetryAfterParserMappingFix()
    Dim answer As VbMsgBoxResult

    answer = MsgBox( _
        "Retry extraction using the parser and mappings currently installed?" & vbCrLf & vbCrLf & _
        "This runs a Full Rescan so unchanged flagged source workbooks are reprocessed. " & _
        "Source DATA is not edited and nothing is approved automatically." & vbCrLf & vbCrLf & _
        "If the same Review Flag returns, the installed parser/mapping still needs correction.", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Retry After Fix")
    If answer <> vbYes Then Exit Sub

    NMDC_FullRescanFast
End Sub


Private Function NMDC_AdminCheckedValue(ByVal value As Variant) As Boolean
    If VarType(value) = vbBoolean Then
        NMDC_AdminCheckedValue = CBool(value)
        Exit Function
    End If

    Select Case UCase$(Trim$(CStr(value)))
        Case "TRUE", "YES", "1", "ON", "CHECKED", "SELECTED"
            NMDC_AdminCheckedValue = True
        Case Else
            NMDC_AdminCheckedValue = False
    End Select
End Function

Private Function NMDC_LatestParserFixRequestPath() As String
    On Error GoTo Missing

    Dim fso As Object
    Dim reportRoot As String
    Dim folder As Object
    Dim folderName As String
    Dim seq As Long
    Dim maxSeq As Long
    Dim candidate As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    reportRoot = NMDC_WorkbookFolder() & "\PARSER_FIX_REPORTS"
    If Not fso.FolderExists(reportRoot) Then Exit Function

    For Each folder In fso.GetFolder(reportRoot).SubFolders
        folderName = CStr(folder.Name)
        If Len(folderName) >= 4 And IsNumeric(folderName) Then
            seq = CLng(folderName)
            If seq > maxSeq Then maxSeq = seq
        End If
    Next folder

    If maxSeq <= 0 Then Exit Function
    candidate = reportRoot & "\" & Format$(maxSeq, "0000") & "\PARSER_FIX_REQUEST.md"
    If fso.FileExists(candidate) Then NMDC_LatestParserFixRequestPath = candidate
    Exit Function

Missing:
    NMDC_LatestParserFixRequestPath = ""
End Function

Private Sub NMDC_RevealFile(ByVal filePath As String)
    On Error Resume Next

    Dim fso As Object
    Dim shell As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    If fso.FileExists(filePath) Then
        shell.Run "explorer.exe /select," & Chr$(34) & filePath & Chr$(34), 1, False
    ElseIf fso.FolderExists(NMDC_WorkbookFolder()) Then
        shell.Run "explorer.exe " & Chr$(34) & NMDC_WorkbookFolder() & Chr$(34), 1, False
    End If
    On Error GoTo 0
End Sub


Private Function NMDC_AdminTableValue(ByVal table As ListObject, ByVal row As ListRow, ByVal headerName As String) As Variant
    On Error GoTo Missing
    NMDC_AdminTableValue = row.Range.Cells(1, table.ListColumns(headerName).Index).Value
    Exit Function
Missing:
    NMDC_AdminTableValue = ""
End Function

Private Function NMDC_AdminCsvField(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    NMDC_AdminCsvField = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Sub NMDC_EnsureAdminFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then Exit Sub
    If Not fso.FolderExists(fso.GetParentFolderName(folderPath)) Then NMDC_EnsureAdminFolder fso.GetParentFolderName(folderPath)
    fso.CreateFolder folderPath
End Sub
