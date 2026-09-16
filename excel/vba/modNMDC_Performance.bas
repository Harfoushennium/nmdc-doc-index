Attribute VB_Name = "modNMDC_Performance"
Option Explicit

Private mAsyncBusy As Boolean
Private mAsyncCommand As String
Private mAsyncCallback As String
Private mAsyncLauncherPath As String
Private mAsyncCompletionPath As String
Private mAsyncStdoutPath As String
Private mAsyncStderrPath As String
Private mAsyncStartedAt As Date
Private mAsyncNextPoll As Date
Private mAsyncPollProcedure As String
Private mAsyncLastExit As Long
Private mAsyncLastDetail As String
Private mAsyncTick As Long

Public Function NMDC_AsyncIsBusy() As Boolean
    NMDC_AsyncIsBusy = mAsyncBusy
End Function

Public Function NMDC_AsyncLastExitCode() As Long
    NMDC_AsyncLastExitCode = mAsyncLastExit
End Function

Public Function NMDC_AsyncLastDetail() As String
    NMDC_AsyncLastDetail = mAsyncLastDetail
End Function

Public Sub NMDC_FastStartup()
    On Error GoTo Handler
    Application.StatusBar = "NMDC Document Index: loading dashboard status..."
    If NMDC_RunEngine("export-excel") = 0 Then NMDC_RefreshDashboardOnlyFast
    NMDC_ApplyOwnerUX
    Application.StatusBar = False
    Exit Sub
Handler:
    Application.StatusBar = False
    NMDC_LogError "FAST_STARTUP_ERROR", _
        "The workbook opened, but the lightweight dashboard startup could not finish.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_UpdateChangedFilesFast()
    NMDC_StartStagingFast False
End Sub

Public Sub NMDC_FullRescanFast()
    NMDC_StartStagingFast True
End Sub

Private Sub NMDC_StartStagingFast(ByVal fullRescan As Boolean)
    On Error GoTo Handler

    Dim dataFolder As String
    Dim modeText As String
    Dim extraArgs As String

    If NMDC_AsyncIsBusy() Then
        MsgBox "A scan is already running." & vbCrLf & vbCrLf & _
               "Excel remains available while it runs. Progress is shown in the Excel status bar.", _
               vbInformation, "NMDC Document Index"
        Exit Sub
    End If

    dataFolder = Trim$(NMDC_ConfigValue("Data Folder"))
    If Len(dataFolder) = 0 Or Not NMDC_FolderExists(dataFolder) Then
        MsgBox "Please select a valid data folder first.", vbExclamation, "NMDC Document Index"
        NMDC_SelectDataFolder
        dataFolder = Trim$(NMDC_ConfigValue("Data Folder"))
        If Len(dataFolder) = 0 Or Not NMDC_FolderExists(dataFolder) Then Exit Sub
    End If

    If Not NMDC_PrepareRulesForStage() Then Exit Sub

    If fullRescan Then
        modeText = "full"
    Else
        modeText = "incremental"
    End If

    extraArgs = "--mode " & modeText & " --data-dir " & NMDC_Quote(dataFolder)
    If Not NMDC_StartEngineAsync("stage", extraArgs, "NMDC_StageFastCompleted") Then Exit Sub

    MsgBox "The scan has started in the background." & vbCrLf & vbCrLf & _
           "Excel should remain usable while the engine works. Progress and elapsed time are shown in the status bar." & vbCrLf & _
           "For normal work use Update Changed Files; Full Rescan is for deliberate rebuilds.", _
           vbInformation, "NMDC Document Index"
    Exit Sub

Handler:
    Application.StatusBar = False
    NMDC_LogError "FAST_STAGE_START_ERROR", _
        "Excel could not start the responsive scan.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_StageFastCompleted()
    On Error GoTo Handler

    Dim exitCode As Long
    exitCode = NMDC_AsyncLastExitCode()
    If exitCode <> 0 Then
        Application.StatusBar = False
        MsgBox "The update could not be completed. Your approved index was not changed." & vbCrLf & vbCrLf & _
               "Please review Error Log for the exact source or technical reason.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    Application.StatusBar = "NMDC Document Index: preparing review tables..."
    exitCode = NMDC_RunEngine("export-excel")
    If exitCode <> 0 Then
        Application.StatusBar = False
        MsgBox "The proposed update was staged, but Excel could not prepare it for review." & vbCrLf & _
               "Your approved index was not changed. Please review Error Log.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    If Not NMDC_RefreshReviewDataFast() Then
        Application.StatusBar = False
        MsgBox "The proposed update was staged, but one or more review tables could not be refreshed." & vbCrLf & _
               "Your approved index was not changed. Please review Error Log.", _
               vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If

    Application.StatusBar = False
    MsgBox "The proposed update is ready for review." & vbCrLf & vbCrLf & _
           "Review Pending Update. Use the native Excel checkboxes in the source panel on the right to include or exclude source workbooks, then Save Source Choices & Restage." & vbCrLf & _
           "When acceptable, return Home and choose Approve Update, Hold Update, or Reject Update.", _
           vbInformation, "NMDC Document Index"
    NMDC_GoToSheet "Pending Update"
    Exit Sub

Handler:
    Application.StatusBar = False
    NMDC_LogError "FAST_STAGE_COMPLETE_ERROR", _
        "Excel could not complete the responsive staged-update workflow.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ReviewPendingUpdateFast()
    If NMDC_RunEngine("export-excel") <> 0 Then
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    If NMDC_RefreshReviewDataFast() Then NMDC_GoToSheet "Pending Update"
End Sub

Public Sub NMDC_ReviewFlagsFast()
    If NMDC_RunEngine("export-excel") <> 0 Then
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    If NMDC_RefreshReviewDataFast() Then NMDC_GoToSheet "Review Flags"
End Sub

Public Sub NMDC_RefreshDashboardFast()
    On Error GoTo Handler
    If NMDC_RunEngine("export-excel") <> 0 Then
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    If Not NMDC_RefreshDashboardOnlyFast() Then
        MsgBox "The dashboard could not be refreshed. Please review Error Log.", _
               vbExclamation, "NMDC Document Index"
        NMDC_GoToSheet "Error Log"
        Exit Sub
    End If
    MsgBox "Dashboard refreshed.", vbInformation, "NMDC Document Index"
    Exit Sub
Handler:
    Application.StatusBar = False
    NMDC_LogError "FAST_DASHBOARD_ERROR", "Excel could not refresh the dashboard.", Err.Number & " - " & Err.Description
End Sub

Public Function NMDC_RefreshReviewDataFast() As Boolean
    On Error GoTo Handler

    Dim ok As Boolean
    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean

    ok = True
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Application.StatusBar = "NMDC Document Index: review refresh 1/5 - Pending Update"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\pending_update.csv", "Pending Update", "PendingUpdate") Then ok = False
    DoEvents

    Application.StatusBar = "NMDC Document Index: review refresh 2/5 - Review Flags"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\flags.csv", "Review Flags", "ReviewFlags") Then ok = False
    DoEvents

    Application.StatusBar = "NMDC Document Index: review refresh 3/5 - Update History"
    If Not NMDC_LoadCsvToTable(NMDC_ExchangePath() & "\history.csv", "Update History", "UpdateHistory") Then ok = False
    DoEvents

    Application.StatusBar = "NMDC Document Index: review refresh 4/5 - Error Log"
    If Not NMDC_LoadErrorCsvPreserveLocal(NMDC_ExchangePath() & "\errors.csv") Then ok = False
    DoEvents

    Application.StatusBar = "NMDC Document Index: review refresh 5/5 - Dashboard / source choices"
    If Not NMDC_LoadDashboard(NMDC_ExchangePath() & "\dashboard.csv") Then ok = False
    NMDC_ApplyOwnerUX
    NMDC_RebuildPendingSourcePanel

CleanExit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    NMDC_RefreshReviewDataFast = ok
    Exit Function
Handler:
    ok = False
    NMDC_LogError "FAST_REVIEW_REFRESH_ERROR", _
        "Excel could not refresh the staged review view.", _
        Err.Number & " - " & Err.Description
    Resume CleanExit
End Function

Public Function NMDC_RefreshDashboardOnlyFast() As Boolean
    On Error GoTo Handler
    Dim ok As Boolean
    ok = True
    Application.StatusBar = "NMDC Document Index: refreshing dashboard..."
    If Not NMDC_LoadErrorCsvPreserveLocal(NMDC_ExchangePath() & "\errors.csv") Then ok = False
    If Not NMDC_LoadDashboard(NMDC_ExchangePath() & "\dashboard.csv") Then ok = False
    NMDC_ApplyOwnerUX
    Application.StatusBar = False
    NMDC_RefreshDashboardOnlyFast = ok
    Exit Function
Handler:
    Application.StatusBar = False
    NMDC_LogError "FAST_DASHBOARD_REFRESH_ERROR", _
        "Excel could not refresh the dashboard view.", Err.Number & " - " & Err.Description
    NMDC_RefreshDashboardOnlyFast = False
End Function

Public Function NMDC_StartEngineAsync(ByVal commandName As String, ByVal extraArgs As String, ByVal callbackMacro As String) As Boolean
    On Error GoTo Handler

    Dim enginePath As String
    Dim runtimePath As String
    Dim exchangePath As String
    Dim commandLine As String
    Dim token As String
    Dim fso As Object
    Dim shell As Object

    If mAsyncBusy Then
        MsgBox "Another NMDC Index operation is already running.", vbInformation, "NMDC Document Index"
        Exit Function
    End If

    enginePath = NMDC_EnginePath()
    runtimePath = NMDC_RuntimePath()
    exchangePath = NMDC_ExchangePath()

    If Not NMDC_FileExists(enginePath) Then
        NMDC_LogError "ENGINE_MISSING", "The NMDC Index engine could not be found.", NMDC_PathDiagnostics()
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(runtimePath) Then fso.CreateFolder runtimePath

    commandLine = NMDC_Quote(enginePath) & " " & commandName & _
                  " --state-dir " & NMDC_Quote(runtimePath) & _
                  " --exchange-dir " & NMDC_Quote(exchangePath) & _
                  " --config-dir " & NMDC_Quote(NMDC_ConfigPath())
    If Len(Trim$(extraArgs)) > 0 Then commandLine = commandLine & " " & extraArgs

    token = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 1000))
    mAsyncLauncherPath = fso.BuildPath(runtimePath, "nmdc_async_" & token & ".cmd")
    mAsyncCompletionPath = fso.BuildPath(runtimePath, "nmdc_async_" & token & ".done")
    mAsyncStdoutPath = fso.BuildPath(runtimePath, "nmdc_async_" & token & ".out.txt")
    mAsyncStderrPath = fso.BuildPath(runtimePath, "nmdc_async_" & token & ".err.txt")

    NMDC_WriteAsyncLauncher mAsyncLauncherPath, mAsyncCompletionPath, mAsyncStdoutPath, mAsyncStderrPath, commandLine

    mAsyncCommand = commandName
    mAsyncCallback = callbackMacro
    mAsyncStartedAt = Now
    mAsyncLastExit = -1
    mAsyncLastDetail = ""
    mAsyncTick = 0
    mAsyncBusy = True
    mAsyncPollProcedure = NMDC_AsyncQualifiedMacro("NMDC_PollEngineAsync")

    Set shell = CreateObject("WScript.Shell")
    shell.Run NMDC_Quote(mAsyncLauncherPath), 0, False
    NMDC_AsyncSchedulePoll
    NMDC_StartEngineAsync = True
    Exit Function

Handler:
    mAsyncBusy = False
    Application.StatusBar = False
    NMDC_LogError "ASYNC_ENGINE_START_ERROR", _
        "Excel could not start the background NMDC Index operation.", _
        Err.Number & " - " & Err.Description & "; " & NMDC_PathDiagnostics()
    NMDC_StartEngineAsync = False
End Function

Public Sub NMDC_PollEngineAsync()
    On Error GoTo Handler

    Dim fso As Object
    Dim callbackName As String
    Dim detail As String

    If Not mAsyncBusy Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    mAsyncTick = mAsyncTick + 1

    If Not fso.FileExists(mAsyncCompletionPath) Then
        NMDC_AsyncShowProgress
        If DateDiff("s", mAsyncStartedAt, Now) > 14400 Then
            Err.Raise vbObjectError + 944, "NMDC Async Engine", "The background engine did not finish within four hours."
        End If
        NMDC_AsyncSchedulePoll
        Exit Sub
    End If

    mAsyncLastExit = CLng(Val(Trim$(NMDC_AsyncReadText(mAsyncCompletionPath))))
    If mAsyncLastExit <> 0 Then
        detail = Trim$(NMDC_AsyncReadText(mAsyncStderrPath))
        If Len(detail) = 0 Then detail = Trim$(NMDC_AsyncReadText(mAsyncStdoutPath))
        If Len(detail) > 4000 Then detail = Left$(detail, 4000)
        mAsyncLastDetail = detail
        NMDC_LogError "ENGINE_EXIT_CODE", _
            "The requested background action did not complete successfully.", _
            "Engine exit code: " & CStr(mAsyncLastExit) & ". Command: " & mAsyncCommand & _
            IIf(Len(detail) > 0, ". Engine message: " & detail, "")
    End If

    callbackName = mAsyncCallback
    mAsyncBusy = False
    Application.StatusBar = False
    NMDC_AsyncCleanup
    If Len(callbackName) > 0 Then Application.Run NMDC_AsyncQualifiedMacro(callbackName)
    Exit Sub

Handler:
    mAsyncBusy = False
    Application.StatusBar = False
    NMDC_AsyncCleanup
    NMDC_LogError "ASYNC_ENGINE_MONITOR_ERROR", _
        "Excel could not monitor the background NMDC Index operation.", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_AsyncSchedulePoll()
    mAsyncNextPoll = Now + TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=mAsyncNextPoll, Procedure:=mAsyncPollProcedure, Schedule:=True
End Sub

Public Sub NMDC_CancelAsyncPoll()
    On Error Resume Next
    If Len(mAsyncPollProcedure) > 0 And mAsyncNextPoll > 0 Then
        Application.OnTime EarliestTime:=mAsyncNextPoll, Procedure:=mAsyncPollProcedure, Schedule:=False
    End If
    On Error GoTo 0
End Sub

Private Function NMDC_AsyncQualifiedMacro(ByVal macroName As String) As String
    ' Use workbook NAME only. Never schedule a SharePoint/OneDrive URL as the macro identity.
    NMDC_AsyncQualifiedMacro = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

Private Sub NMDC_AsyncShowProgress()
    Dim elapsed As Long
    Dim position As Long
    Dim bar As String
    position = (mAsyncTick Mod 18) + 1
    bar = "[" & String$(position - 1, ChrW(183)) & String$(4, ChrW(9632)) & String$(18 - position, ChrW(183)) & "]"
    elapsed = DateDiff("s", mAsyncStartedAt, Now)
    Application.StatusBar = "NMDC Document Index - scanning in background " & bar & _
                            "  Elapsed: " & CStr(elapsed) & " s  | Excel remains available"
End Sub

Private Sub NMDC_WriteAsyncLauncher(ByVal launcherPath As String, ByVal completionPath As String, _
                                    ByVal stdoutPath As String, ByVal stderrPath As String, _
                                    ByVal commandLine As String)
    Dim fso As Object
    Dim file As Object
    Dim safeCommand As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    safeCommand = Replace(commandLine, "%", "%%")
    Set file = fso.OpenTextFile(launcherPath, 2, True)
    file.WriteLine "@echo off"
    file.WriteLine safeCommand & " 1> " & NMDC_Quote(stdoutPath) & " 2> " & NMDC_Quote(stderrPath)
    file.WriteLine "set ""NMDC_EXIT=%ERRORLEVEL%"""
    file.WriteLine "> " & NMDC_Quote(completionPath) & " echo %NMDC_EXIT%"
    file.WriteLine "exit /b %NMDC_EXIT%"
    file.Close
End Sub

Private Function NMDC_AsyncReadText(ByVal filePath As String) As String
    On Error GoTo Missing
    Dim fso As Object
    Dim file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    Set file = fso.OpenTextFile(filePath, 1, False)
    NMDC_AsyncReadText = file.ReadAll
    file.Close
    Exit Function
Missing:
    NMDC_AsyncReadText = ""
End Function

Private Sub NMDC_AsyncCleanup()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(mAsyncLauncherPath) Then fso.DeleteFile mAsyncLauncherPath, True
    If fso.FileExists(mAsyncCompletionPath) Then fso.DeleteFile mAsyncCompletionPath, True
    If fso.FileExists(mAsyncStdoutPath) Then fso.DeleteFile mAsyncStdoutPath, True
    If fso.FileExists(mAsyncStderrPath) Then fso.DeleteFile mAsyncStderrPath, True
    mAsyncPollProcedure = ""
    mAsyncNextPoll = 0
    On Error GoTo 0
End Sub
