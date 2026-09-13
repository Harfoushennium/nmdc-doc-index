Attribute VB_Name = "modNMDC_Actions"
Option Explicit

Public Sub NMDC_UpdateChangedFiles()
    NMDC_RunStagingAction False
End Sub

Public Sub NMDC_FullRescan()
    NMDC_RunStagingAction True
End Sub

Private Sub NMDC_RunStagingAction(ByVal fullRescan As Boolean)
    On Error GoTo Handler

    Dim dataFolder As String
    Dim modeText As String
    Dim exitCode As Long

    dataFolder = Trim$(NMDC_ConfigValue("Data Folder"))
    If Len(dataFolder) = 0 Or Len(Dir$(dataFolder, vbDirectory)) = 0 Then
        MsgBox "Please select a valid data folder first.", vbExclamation, "NMDC Document Index"
        NMDC_SelectDataFolder
        dataFolder = Trim$(NMDC_ConfigValue("Data Folder"))
        If Len(dataFolder) = 0 Or Len(Dir$(dataFolder, vbDirectory)) = 0 Then Exit Sub
    End If

    If fullRescan Then
        modeText = "full"
    Else
        modeText = "incremental"
    End If

    Application.StatusBar = "NMDC Document Index: preparing staged update..."
    exitCode = NMDC_RunEngine("stage", "--mode " & modeText & " --data-dir " & NMDC_Quote(dataFolder))
    If exitCode <> 0 Then
        Application.StatusBar = False
        MsgBox "The update could not be completed. Your approved index was not changed." & vbCrLf & _
               "Please review the Error Log.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    exitCode = NMDC_RunEngine("export-excel")
    If exitCode <> 0 Then
        Application.StatusBar = False
        MsgBox "The proposed update was staged, but Excel could not load it for review." & vbCrLf & vbCrLf & _
               "Your approved index was not changed. Do not approve until Refresh Dashboard succeeds.", _
               vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If

    If Not NMDC_RefreshExchangeData() Then
        Application.StatusBar = False
        MsgBox "The proposed update was staged, but one or more review tables could not be refreshed." & vbCrLf & vbCrLf & _
               "Your approved index was not changed. Do not approve until Refresh Dashboard succeeds.", _
               vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    Application.StatusBar = False

    MsgBox "The proposed update is ready for your review." & vbCrLf & vbCrLf & _
           "Please check Pending Update and Review Flags before approving.", _
           vbInformation, "NMDC Document Index"
    NMDC_GoToSheet "Pending Update"
    Exit Sub

Handler:
    Application.StatusBar = False
    NMDC_LogError "STAGE_ACTION_ERROR", "Excel could not prepare the staged update.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ApproveUpdate()
    On Error GoTo Handler

    Dim answer As VbMsgBoxResult
    Dim exitCode As Long
    Dim reviewedRunId As String
    Dim currentRunId As String
    Dim currentStatus As String
    Dim conflictCount As Long
    Dim approvalArgs As String

    reviewedRunId = NMDC_DisplayedPendingRunId()
    If Len(reviewedRunId) = 0 Then
        MsgBox "There is no displayed pending update to approve." & vbCrLf & _
               "Please press Review Pending Update first.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    exitCode = NMDC_RunEngine("export-excel")
    If exitCode <> 0 Then
        MsgBox "Excel could not verify that the displayed proposal is still current." & vbCrLf & vbCrLf & _
               "Nothing was approved. Please review the Error Log.", vbCritical, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    If Not NMDC_RefreshExchangeData() Then
        MsgBox "Excel could not verify that the displayed proposal is still current." & vbCrLf & vbCrLf & _
               "Nothing was approved. Please review the Error Log.", vbCritical, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If

    currentRunId = NMDC_DisplayedPendingRunId()
    If Len(currentRunId) = 0 Or StrComp(reviewedRunId, currentRunId, vbBinaryCompare) <> 0 Then
        MsgBox "The pending update changed after your last review." & vbCrLf & vbCrLf & _
               "Nothing was approved. Please review the current Pending Update before approving.", _
               vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Pending Update"
        Exit Sub
    End If

    currentStatus = UCase$(Trim$(CStr(ThisWorkbook.Worksheets("Home").Range("H5").Value)))
    If currentStatus <> "STAGED" And currentStatus <> "REVIEW_REQUIRED" And currentStatus <> "HOLD" Then
        MsgBox "Update " & reviewedRunId & " is not awaiting approval (status: " & currentStatus & ")." & vbCrLf & vbCrLf & _
               "Nothing was approved. Please create or review a pending update first.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    conflictCount = CLng(Val(ThisWorkbook.Worksheets("Home").Range("K6").Value))
    If conflictCount > 0 Then
        answer = MsgBox("This staged update has " & CStr(conflictCount) & " conflict flag(s)." & vbCrLf & vbCrLf & _
                        "Continue only if you reviewed and intentionally accept every overridable conflict." & vbCrLf & _
                        "Parser, source-hash and duplicate-key conflicts can never be overridden." & vbCrLf & vbCrLf & _
                        "Attempt approval of update " & reviewedRunId & "?", _
                        vbExclamation + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    Else
        answer = MsgBox("Approve staged update " & reviewedRunId & _
                        " and make it the new approved master index?" & vbCrLf & vbCrLf & _
                        "This action is recorded in the update history.", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    End If
    If answer <> vbYes Then Exit Sub

    approvalArgs = "--run-id " & NMDC_Quote(reviewedRunId)
    If conflictCount > 0 Then approvalArgs = approvalArgs & " --allow-conflicts"
    exitCode = NMDC_RunEngine("approve", approvalArgs)
    If exitCode <> 0 Then
        MsgBox "The staged update could not be approved. The previous approved index remains unchanged." & vbCrLf & _
               "Please review Review Flags and Error Log.", vbCritical, "NMDC Document Index"
        Exit Sub
    End If

    exitCode = NMDC_RunEngine("export-excel")
    If exitCode <> 0 Then
        MsgBox "Update " & reviewedRunId & " was approved safely, but Excel could not refresh the displayed tables." & vbCrLf & vbCrLf & _
               "Please review the Error Log and press Refresh Dashboard.", vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    If Not NMDC_RefreshExchangeData() Then
        MsgBox "Update " & reviewedRunId & " was approved safely, but Excel could not refresh the displayed tables." & vbCrLf & vbCrLf & _
               "Please review the Error Log and press Refresh Dashboard.", vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    MsgBox "Update approved successfully.", vbInformation, "NMDC Document Index"
    NMDC_GoToSheet "Home"
    Exit Sub

Handler:
    NMDC_LogError "APPROVE_ACTION_ERROR", "Excel could not safely approve the staged update.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_HoldUpdate()
    Dim note As String
    Dim runId As String

    runId = NMDC_DisplayedPendingRunId()
    If Len(runId) = 0 Then
        MsgBox "There is no displayed pending update to hold.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If
    note = InputBox("Optional note for holding this update:", "Hold Update")
    If NMDC_RunEngine("hold", "--run-id " & NMDC_Quote(runId) & " --note " & NMDC_Quote(note)) = 0 Then
        If NMDC_RunEngine("export-excel") <> 0 Then
            MsgBox "The update was placed on hold, but Excel could not refresh the displayed tables." & vbCrLf & _
                   "Please review the Error Log.", vbExclamation, "NMDC Document Index"
            NMDC_GoToSheet "Error Log"
            Exit Sub
        End If
        If Not NMDC_RefreshExchangeData() Then
            MsgBox "The update was placed on hold, but Excel could not refresh the displayed tables." & vbCrLf & _
                   "Please review the Error Log.", vbExclamation, "NMDC Document Index"
            NMDC_GoToSheet "Error Log"
            Exit Sub
        End If
        MsgBox "The update is on hold. The approved master index was not changed.", vbInformation, "NMDC Document Index"
    End If
End Sub

Public Sub NMDC_RejectUpdate()
    Dim answer As VbMsgBoxResult
    Dim note As String
    Dim runId As String

    runId = NMDC_DisplayedPendingRunId()
    If Len(runId) = 0 Then
        MsgBox "There is no displayed pending update to reject.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    answer = MsgBox("Reject the current staged update?" & vbCrLf & _
                    "The approved master index will remain unchanged.", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer <> vbYes Then Exit Sub

    note = InputBox("Optional reason for rejection:", "Reject Update")
    If NMDC_RunEngine("reject", "--run-id " & NMDC_Quote(runId) & " --note " & NMDC_Quote(note)) = 0 Then
        If NMDC_RunEngine("export-excel") <> 0 Then
            MsgBox "The update was rejected, but Excel could not refresh the displayed tables." & vbCrLf & _
                   "Please review the Error Log.", vbExclamation, "NMDC Document Index"
            NMDC_GoToSheet "Error Log"
            Exit Sub
        End If
        If Not NMDC_RefreshExchangeData() Then
            MsgBox "The update was rejected, but Excel could not refresh the displayed tables." & vbCrLf & _
                   "Please review the Error Log.", vbExclamation, "NMDC Document Index"
            NMDC_GoToSheet "Error Log"
            Exit Sub
        End If
        MsgBox "The staged update was rejected. The approved master index was not changed.", vbInformation, "NMDC Document Index"
    End If
End Sub

Public Sub NMDC_SelectDataFolder()
    On Error GoTo Handler

    Dim picker As FileDialog
    Dim selectedPath As String

    Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    picker.Title = "Select the folder containing NMDC document registers"
    picker.AllowMultiSelect = False

    If picker.Show <> -1 Then Exit Sub
    selectedPath = picker.SelectedItems(1)
    NMDC_SetConfigValue "Data Folder", selectedPath
    ThisWorkbook.Worksheets("Home").Range("B7").Value = selectedPath
    MsgBox "Data folder saved." & vbCrLf & selectedPath, vbInformation, "NMDC Document Index"
    Exit Sub

Handler:
    NMDC_LogError "FOLDER_PICKER_ERROR", "Excel could not select the data folder.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_RefreshDashboard()
    If NMDC_RunEngine("export-excel") = 0 Then
        If NMDC_RefreshExchangeData() Then
            MsgBox "Dashboard refreshed.", vbInformation, "NMDC Document Index"
        Else
            MsgBox "Excel could not refresh all index tables. Please review the Error Log.", _
                   vbExclamation, "NMDC Document Index"
            NMDC_GoToSheet "Error Log"
        End If
    End If
End Sub

Public Sub NMDC_ReviewPendingUpdate()
    NMDC_RefreshAndOpen "Pending Update"
End Sub

Public Sub NMDC_ReviewFlags()
    NMDC_RefreshAndOpen "Review Flags"
End Sub

Public Sub NMDC_OpenConfiguration()
    NMDC_GoToSheet "Configuration"
End Sub

Public Sub NMDC_OpenRulesMappings()
    NMDC_GoToSheet "Rules & Mappings"
End Sub

Public Sub NMDC_OpenHelp()
    NMDC_GoToSheet "Help"
End Sub

Public Sub NMDC_ViewLog()
    Dim refreshSucceeded As Boolean
    If NMDC_RunEngine("export-excel") = 0 Then refreshSucceeded = NMDC_RefreshExchangeData()
    NMDC_GoToSheet "Error Log"
End Sub

Public Sub NMDC_FlagWrongData()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim projectNo As String
    Dim documentNo As String
    Dim sourceFile As String
    Dim sourceSheet As String
    Dim sourceRow As String
    Dim sourceCell As String
    Dim revision As String
    Dim eventIdentity As String
    Dim flagCode As String
    Dim currentValue As String
    Dim currentField As String
    Dim userNote As String
    Dim expectedValue As String
    Dim extraArgs As String

    Set ws = ActiveSheet
    rowNumber = ActiveCell.Row
    If rowNumber <= 1 Then
        MsgBox "Please select a data row first.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    projectNo = NMDC_RowValueByHeader(ws, rowNumber, "Project No.")
    documentNo = NMDC_RowValueByHeader(ws, rowNumber, "Document No.")
    sourceFile = NMDC_RowValueByHeader(ws, rowNumber, "Source File")
    sourceSheet = NMDC_RowValueByHeader(ws, rowNumber, "Source Sheet")
    sourceRow = NMDC_RowValueByHeader(ws, rowNumber, "Source Row")
    sourceCell = NMDC_RowValueByHeader(ws, rowNumber, "Source Cell")
    revision = NMDC_RowValueByHeader(ws, rowNumber, "Revision")
    eventIdentity = NMDC_RowValueByHeader(ws, rowNumber, "Event Key")
    If Len(eventIdentity) = 0 Then eventIdentity = NMDC_RowValueByHeader(ws, rowNumber, "Record Identity")
    flagCode = NMDC_RowValueByHeader(ws, rowNumber, "Flag Code")
    currentField = CStr(ws.Cells(1, ActiveCell.Column).Value)
    currentValue = CStr(ActiveCell.Value)
    userNote = InputBox("Describe what is wrong with this record:", "Flag Wrong Data")
    If Len(Trim$(userNote)) = 0 Then Exit Sub
    expectedValue = InputBox("Enter the correct/expected value if known (optional):", "Flag Wrong Data")

    extraArgs = "--message " & NMDC_Quote(userNote) & _
                " --project-no " & NMDC_Quote(projectNo) & _
                " --document-no " & NMDC_Quote(documentNo) & _
                " --source-file " & NMDC_Quote(sourceFile) & _
                " --worksheet " & NMDC_Quote(sourceSheet) & _
                " --source-row " & NMDC_Quote(sourceRow) & _
                " --source-cell " & NMDC_Quote(sourceCell) & _
                " --revision " & NMDC_Quote(revision) & _
                " --event-identity " & NMDC_Quote(eventIdentity) & _
                " --flag-code " & NMDC_Quote(flagCode) & _
                " --current-field " & NMDC_Quote(currentField) & _
                " --current-value " & NMDC_Quote(currentValue) & _
                " --expected-value " & NMDC_Quote(expectedValue) & _
                " --user-name " & NMDC_Quote(Application.UserName)

    If NMDC_RunEngine("user-flag", extraArgs) = 0 Then
        MsgBox "Your flag has been recorded for review.", vbInformation, "NMDC Document Index"
    End If
    Exit Sub

Handler:
    NMDC_LogError "USER_FLAG_ERROR", "Excel could not record the user flag.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ReportRequirement()
    On Error GoTo Handler

    Dim message As String
    Dim contextArgs As String
    Dim ws As Worksheet
    Dim rowNumber As Long

    message = InputBox("Describe the problem or new requirement in plain English:", "Report Requirement / Problem")
    If Len(Trim$(message)) = 0 Then Exit Sub

    Set ws = ActiveSheet
    rowNumber = ActiveCell.Row
    contextArgs = " --user-name " & NMDC_Quote(Application.UserName)
    If rowNumber > 1 Then
        contextArgs = contextArgs & _
                      " --project-no " & NMDC_Quote(NMDC_RowValueByHeader(ws, rowNumber, "Project No.")) & _
                      " --document-no " & NMDC_Quote(NMDC_RowValueByHeader(ws, rowNumber, "Document No.")) & _
                      " --revision " & NMDC_Quote(NMDC_RowValueByHeader(ws, rowNumber, "Revision")) & _
                      " --source-file " & NMDC_Quote(NMDC_RowValueByHeader(ws, rowNumber, "Source File")) & _
                      " --worksheet " & NMDC_Quote(NMDC_RowValueByHeader(ws, rowNumber, "Source Sheet")) & _
                      " --source-row " & NMDC_Quote(NMDC_RowValueByHeader(ws, rowNumber, "Source Row")) & _
                      " --source-cell " & NMDC_Quote(NMDC_RowValueByHeader(ws, rowNumber, "Source Cell")) & _
                      " --event-identity " & NMDC_Quote(NMDC_RowValueByHeader(ws, rowNumber, "Event Key")) & _
                      " --current-field " & NMDC_Quote(CStr(ws.Cells(1, ActiveCell.Column).Value)) & _
                      " --current-value " & NMDC_Quote(CStr(ActiveCell.Value))
    End If

    If NMDC_RunEngine("support-request", "--message " & NMDC_Quote(message) & contextArgs) = 0 Then
        MsgBox "A support package was created successfully." & vbCrLf & _
               "You can provide that package to GPT/developer support.", vbInformation, "NMDC Document Index"
    End If
    Exit Sub

Handler:
    NMDC_LogError "SUPPORT_REQUEST_ERROR", "Excel could not create the support request.", Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_RefreshAndOpen(ByVal sheetName As String)
    If NMDC_RunEngine("export-excel") <> 0 Then
        MsgBox "Excel could not refresh " & sheetName & ". The previous view may be out of date." & vbCrLf & _
               "Please review the Error Log.", vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    If Not NMDC_RefreshExchangeData() Then
        MsgBox "Excel could not safely refresh " & sheetName & "." & vbCrLf & _
               "Please review the Error Log.", vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    NMDC_GoToSheet sheetName
End Sub

Private Function NMDC_DisplayedPendingRunId() As String
    On Error GoTo Handler
    NMDC_DisplayedPendingRunId = Trim$(CStr(ThisWorkbook.Worksheets("Home").Range("H6").Value))
    Exit Function
Handler:
    NMDC_DisplayedPendingRunId = ""
End Function

Public Sub NMDC_GoToSheet(ByVal sheetName As String)
    On Error GoTo Handler
    ThisWorkbook.Worksheets(sheetName).Activate
    Exit Sub
Handler:
    NMDC_LogError "NAVIGATION_ERROR", "Excel could not open " & sheetName & ".", Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_RowValueByHeader(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal headerName As String) As String
    Dim hit As Range
    Set hit = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If hit Is Nothing Then
        NMDC_RowValueByHeader = ""
    Else
        NMDC_RowValueByHeader = CStr(ws.Cells(rowNumber, hit.Column).Value)
    End If
End Function
