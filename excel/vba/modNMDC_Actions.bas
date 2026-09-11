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
    If exitCode = 0 Then NMDC_RefreshExchangeData
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
    Dim answer As VbMsgBoxResult
    Dim exitCode As Long

    answer = MsgBox("Approve the current staged update and make it the new approved master index?" & vbCrLf & vbCrLf & _
                    "This action is recorded in the update history.", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer <> vbYes Then Exit Sub

    exitCode = NMDC_RunEngine("approve")
    If exitCode <> 0 Then
        MsgBox "The staged update could not be approved. The previous approved index remains unchanged." & vbCrLf & _
               "Please review Review Flags and Error Log.", vbCritical, "NMDC Document Index"
        Exit Sub
    End If

    NMDC_RunEngine "export-excel"
    NMDC_RefreshExchangeData
    MsgBox "Update approved successfully.", vbInformation, "NMDC Document Index"
    NMDC_GoToSheet "Home"
End Sub

Public Sub NMDC_HoldUpdate()
    Dim note As String
    note = InputBox("Optional note for holding this update:", "Hold Update")
    If NMDC_RunEngine("hold", "--note " & NMDC_Quote(note)) = 0 Then
        NMDC_RunEngine "export-excel"
        NMDC_RefreshExchangeData
        MsgBox "The update is on hold. The approved master index was not changed.", vbInformation, "NMDC Document Index"
    End If
End Sub

Public Sub NMDC_RejectUpdate()
    Dim answer As VbMsgBoxResult
    Dim note As String

    answer = MsgBox("Reject the current staged update?" & vbCrLf & _
                    "The approved master index will remain unchanged.", _
                    vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index")
    If answer <> vbYes Then Exit Sub

    note = InputBox("Optional reason for rejection:", "Reject Update")
    If NMDC_RunEngine("reject", "--note " & NMDC_Quote(note)) = 0 Then
        NMDC_RunEngine "export-excel"
        NMDC_RefreshExchangeData
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
        NMDC_RefreshExchangeData
        MsgBox "Dashboard refreshed.", vbInformation, "NMDC Document Index"
    End If
End Sub

Public Sub NMDC_ReviewPendingUpdate()
    NMDC_GoToSheet "Pending Update"
End Sub

Public Sub NMDC_ReviewFlags()
    NMDC_GoToSheet "Review Flags"
End Sub

Public Sub NMDC_ViewLog()
    NMDC_GoToSheet "Error Log"
End Sub

Public Sub NMDC_FlagWrongData()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim documentNo As String
    Dim sourceFile As String
    Dim userNote As String
    Dim expectedValue As String
    Dim extraArgs As String

    Set ws = ActiveSheet
    rowNumber = ActiveCell.Row
    If rowNumber <= 1 Then
        MsgBox "Please select a data row first.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    documentNo = NMDC_RowValueByHeader(ws, rowNumber, "Document No.")
    sourceFile = NMDC_RowValueByHeader(ws, rowNumber, "Source File")
    userNote = InputBox("Describe what is wrong with this record:", "Flag Wrong Data")
    If Len(Trim$(userNote)) = 0 Then Exit Sub
    expectedValue = InputBox("Enter the correct/expected value if known (optional):", "Flag Wrong Data")

    extraArgs = "--message " & NMDC_Quote(userNote) & _
                " --document-no " & NMDC_Quote(documentNo) & _
                " --source-file " & NMDC_Quote(sourceFile) & _
                " --expected-value " & NMDC_Quote(expectedValue)

    If NMDC_RunEngine("user-flag", extraArgs) = 0 Then
        MsgBox "Your flag has been recorded for review.", vbInformation, "NMDC Document Index"
    End If
    Exit Sub

Handler:
    NMDC_LogError "USER_FLAG_ERROR", "Excel could not record the user flag.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ReportRequirement()
    Dim message As String
    message = InputBox("Describe the problem or new requirement in plain English:", "Report Requirement / Problem")
    If Len(Trim$(message)) = 0 Then Exit Sub

    If NMDC_RunEngine("support-request", "--message " & NMDC_Quote(message)) = 0 Then
        MsgBox "A support package was created successfully." & vbCrLf & _
               "You can provide that package to GPT/developer support.", vbInformation, "NMDC Document Index"
    End If
End Sub

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
