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
    Dim exitCode As Long

    Set ws = ThisWorkbook.Worksheets("Review Flags")
    Set table = ws.ListObjects("ReviewFlags")
    filePath = NMDC_ExchangePath() & "\review_decisions.csv"
    NMDC_EnsureAdminFolder NMDC_ExchangePath()

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText "Flag Code,Source File,Source Sheet,Event Key,Project No.,Document No.,Revision,User Decision,User Comment,Resolution Status" & vbLf

    For Each row In table.ListRows
        decision = Trim$(CStr(NMDC_AdminTableValue(table, row, "User Decision")))
        comment = Trim$(CStr(NMDC_AdminTableValue(table, row, "User Comment")))
        status = UCase$(Trim$(CStr(NMDC_AdminTableValue(table, row, "Resolution Status"))))

        If Len(decision) > 0 Or Len(comment) > 0 Or (Len(status) > 0 And status <> "OPEN") Then
            stream.WriteText _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Flag Code"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Source File"))) & "," & _
                NMDC_AdminCsvField(CStr(NMDC_AdminTableValue(table, row, "Source Sheet"))) & "," & _
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

    exitCode = NMDC_RunEngine("save-review-decisions", "--decisions-file " & NMDC_Quote(filePath))
    If exitCode <> 0 Then
        MsgBox "Review decisions could not be saved. Please review the Error Log.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    If NMDC_RunEngine("export-excel") = 0 Then NMDC_RefreshExchangeData
    MsgBox CStr(savedRows) & " Review Flag decision(s) were saved to the staged update audit trail.", _
           vbInformation, "NMDC Document Index"
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
