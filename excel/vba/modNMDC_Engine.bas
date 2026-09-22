Attribute VB_Name = "modNMDC_Engine"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const ENGINE_RELATIVE_PATH As String = "engine\nmdc_index_engine.exe"
Private Const RUNTIME_RELATIVE_PATH As String = "runtime"
Private Const EXCHANGE_RELATIVE_PATH As String = "excel_exchange"
Private Const CONFIG_RELATIVE_PATH As String = "config"

Public Function NMDC_WorkbookFolder() As String
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    If Left$(wbPath, 4) = "http" Then
        Dim localBase As String
        localBase = NMDC_ResolveOneDriveLocalBase(wbPath)
        If Len(localBase) > 0 Then
            NMDC_WorkbookFolder = localBase
            Exit Function
        End If
    End If
    NMDC_WorkbookFolder = wbPath
End Function

Private Function NMDC_ResolveOneDriveLocalBase(ByVal webUrl As String) As String
    On Error GoTo Fallback
    Dim shell As Object
    Dim fso As Object
    Dim userFolder As String
    Dim envVal As String
    Dim relPath As String
    Dim marker As String
    Dim pos As Long
    
    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check OneDrive Commercial Account
    On Error Resume Next
    userFolder = shell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\Business1\UserFolder")
    On Error GoTo Fallback
    
    If Len(userFolder) > 0 And fso.FolderExists(userFolder) Then
        marker = "NMDC DOCUMENTS INDEX"
        pos = InStr(1, webUrl, marker, vbTextCompare)
        If pos > 0 Then
            relPath = Mid$(webUrl, pos)
            relPath = Replace(relPath, "/", "\")
            relPath = Replace(relPath, "%20", " ")
            ' Try building from Desktop/NPCC/AI PROJECTS
            Dim candidate As String
            candidate = fso.BuildPath(userFolder, "Desktop\NPCC\AI PROJECTS\" & relPath)
            If fso.FolderExists(candidate) Then
                NMDC_ResolveOneDriveLocalBase = candidate
                Exit Function
            End If
            candidate = fso.BuildPath(userFolder, relPath)
            If fso.FolderExists(candidate) Then
                NMDC_ResolveOneDriveLocalBase = candidate
                Exit Function
            End If
        End If
    End If

Fallback:
    NMDC_ResolveOneDriveLocalBase = ""
End Function

Public Function NMDC_EnginePath() As String
    Dim configured As String
    configured = Trim$(NMDC_ConfigValue("Engine Executable Path"))
    If Len(configured) > 0 Then
        NMDC_EnginePath = NMDC_ResolvePackagePath(configured)
    Else
        NMDC_EnginePath = NMDC_WorkbookFolder() & "\" & ENGINE_RELATIVE_PATH
    End If
End Function

Public Function NMDC_RuntimePath() As String
    Dim configured As String
    configured = Trim$(NMDC_ConfigValue("Runtime Folder"))
    If Len(configured) > 0 Then
        NMDC_RuntimePath = NMDC_ResolvePackagePath(configured)
    Else
        NMDC_RuntimePath = NMDC_WorkbookFolder() & "\" & RUNTIME_RELATIVE_PATH
    End If
End Function

Public Function NMDC_ExchangePath() As String
    NMDC_ExchangePath = NMDC_RuntimePath() & "\" & EXCHANGE_RELATIVE_PATH
End Function

Public Function NMDC_ConfigPath() As String
    Dim configured As String
    configured = Trim$(NMDC_ConfigValue("Configuration Folder"))
    If Len(configured) > 0 Then
        NMDC_ConfigPath = NMDC_ResolvePackagePath(configured)
    Else
        NMDC_ConfigPath = NMDC_WorkbookFolder() & "\" & CONFIG_RELATIVE_PATH
    End If
End Function

Public Function NMDC_ResolvePackagePath(ByVal configuredPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(fso.GetDriveName(configuredPath)) > 0 Or Left$(configuredPath, 2) = "\\" Then
        NMDC_ResolvePackagePath = configuredPath
    Else
        NMDC_ResolvePackagePath = fso.BuildPath(NMDC_WorkbookFolder(), configuredPath)
    End If
End Function

Public Function NMDC_FileExists(ByVal filePath As String) As Boolean
    On Error GoTo Missing
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    NMDC_FileExists = fso.FileExists(filePath)
    Exit Function
Missing:
    NMDC_FileExists = False
End Function

Public Function NMDC_FolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo Missing
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    NMDC_FolderExists = fso.FolderExists(folderPath)
    Exit Function
Missing:
    NMDC_FolderExists = False
End Function

Public Function NMDC_PathDiagnostics() As String
    NMDC_PathDiagnostics = "Workbook=" & NMDC_WorkbookFolder() & _
        "; Engine=" & NMDC_EnginePath() & _
        "; Runtime=" & NMDC_RuntimePath() & _
        "; Config=" & NMDC_ConfigPath()
End Function

Public Function NMDC_Quote(ByVal value As String) As String
    NMDC_Quote = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Public Function NMDC_RunEngine(ByVal commandName As String, Optional ByVal extraArgs As String = "") As Long
    On Error GoTo Handler

    Dim enginePath As String
    Dim runtimePath As String
    Dim exchangePath As String
    Dim cmd As String
    Dim shell As Object
    Dim fso As Object
    Dim exitCode As Long
    Dim startedAt As Date
    Dim tick As Long
    Dim runToken As String
    Dim launcherPath As String
    Dim completionPath As String
    Dim stdoutPath As String
    Dim stderrPath As String
    Dim engineDetail As String

    enginePath = NMDC_EnginePath()
    runtimePath = NMDC_RuntimePath()
    exchangePath = NMDC_ExchangePath()

    If Not NMDC_FileExists(enginePath) Then
        NMDC_LogError "ENGINE_MISSING", _
            "The NMDC Index engine could not be found.", _
            NMDC_PathDiagnostics()
        MsgBox "The NMDC Index engine could not be found." & vbCrLf & vbCrLf & _
               "Nothing was changed. Please open Error Log or use Report Requirement / Problem.", _
               vbExclamation, "NMDC Document Index"
        NMDC_RunEngine = 9001
        Exit Function
    End If

    If Not NMDC_FolderExists(NMDC_ConfigPath()) Then
        NMDC_LogError "CONFIG_FOLDER_MISSING", _
            "The NMDC Index configuration folder could not be found.", _
            NMDC_PathDiagnostics()
        MsgBox "The NMDC Index configuration folder could not be found." & vbCrLf & vbCrLf & _
               "Nothing was changed. Please rerun the one-time setup from the extracted production package.", _
               vbExclamation, "NMDC Document Index"
        NMDC_RunEngine = 9003
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(runtimePath) Then fso.CreateFolder runtimePath

    cmd = NMDC_Quote(enginePath) & " " & commandName & _
          " --state-dir " & NMDC_Quote(runtimePath) & _
          " --exchange-dir " & NMDC_Quote(exchangePath) & _
          " --config-dir " & NMDC_Quote(NMDC_ConfigPath())
    If Len(Trim$(extraArgs)) > 0 Then cmd = cmd & " " & extraArgs

    runToken = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 1000))
    launcherPath = fso.BuildPath(runtimePath, "nmdc_engine_" & runToken & ".cmd")
    completionPath = fso.BuildPath(runtimePath, "nmdc_engine_" & runToken & ".done")
    stdoutPath = fso.BuildPath(runtimePath, "nmdc_engine_" & runToken & ".out.txt")
    stderrPath = fso.BuildPath(runtimePath, "nmdc_engine_" & runToken & ".err.txt")

    NMDC_WriteHiddenLauncher launcherPath, completionPath, stdoutPath, stderrPath, cmd

    Set shell = CreateObject("WScript.Shell")
    startedAt = Now
    tick = 0
    Application.Cursor = xlWait

    ' Run the generated command script hidden and return immediately. Excel then
    ' polls the completion marker, so the UI stays responsive with no black console window.
    shell.Run NMDC_Quote(launcherPath), 0, False

    Do While Not fso.FileExists(completionPath)
        tick = tick + 1
        NMDC_ShowEngineProgress commandName, startedAt, tick
        DoEvents
        Sleep 140
        If DateDiff("s", startedAt, Now) > 14400 Then
            Err.Raise vbObjectError + 904, "NMDC Engine", "The engine did not finish within four hours."
        End If
    Loop

    exitCode = CLng(Val(Trim$(NMDC_ReadTextFile(completionPath))))
    If exitCode <> 0 Then
        engineDetail = Trim$(NMDC_ReadTextFile(stderrPath))
        If Len(engineDetail) = 0 Then engineDetail = Trim$(NMDC_ReadTextFile(stdoutPath))
        If Len(engineDetail) > 4000 Then engineDetail = Left$(engineDetail, 4000)
    End If

    Application.Cursor = xlDefault
    Application.StatusBar = False
    NMDC_RunEngine = exitCode

    If exitCode <> 0 Then
        NMDC_LogError "ENGINE_EXIT_CODE", _
            "The requested action did not complete successfully.", _
            "Engine exit code: " & CStr(exitCode) & ". Command: " & commandName & _
            IIf(Len(engineDetail) > 0, ". Engine message: " & engineDetail, "")
    End If

    NMDC_DeleteIfExists launcherPath
    NMDC_DeleteIfExists completionPath
    NMDC_DeleteIfExists stdoutPath
    NMDC_DeleteIfExists stderrPath
    Exit Function

Handler:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    NMDC_DeleteIfExists launcherPath
    NMDC_DeleteIfExists completionPath
    NMDC_DeleteIfExists stdoutPath
    NMDC_DeleteIfExists stderrPath
    NMDC_LogError "VBA_ENGINE_LAUNCH_ERROR", _
        "Excel could not start or monitor the NMDC Index engine.", _
        Err.Number & " - " & Err.Description & "; " & NMDC_PathDiagnostics()
    NMDC_RunEngine = 9002
End Function

Private Sub NMDC_WriteHiddenLauncher(ByVal launcherPath As String, ByVal completionPath As String, _
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

Private Function NMDC_ReadTextFile(ByVal filePath As String) As String
    On Error GoTo Missing
    Dim fso As Object
    Dim file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    Set file = fso.OpenTextFile(filePath, 1, False)
    NMDC_ReadTextFile = file.ReadAll
    file.Close
    Exit Function
Missing:
    NMDC_ReadTextFile = ""
End Function

Private Sub NMDC_DeleteIfExists(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
End Sub

Private Sub NMDC_ShowEngineProgress(ByVal commandName As String, ByVal startedAt As Date, ByVal tick As Long)
    Dim phaseText As String
    Dim elapsedSeconds As Long
    Dim barWidth As Long
    Dim blockWidth As Long
    Dim position As Long
    Dim progressBar As String

    Select Case LCase$(commandName)
        Case "stage"
            phaseText = "Scanning and analysing source workbooks"
        Case "export-excel"
            phaseText = "Preparing Excel review tables"
        Case "approve"
            phaseText = "Approving staged update"
        Case "hold"
            phaseText = "Placing staged update on hold"
        Case "reject"
            phaseText = "Rejecting staged update"
        Case "user-flag"
            phaseText = "Saving user review flag"
        Case Else
            phaseText = "Running " & commandName
    End Select

    barWidth = 22
    blockWidth = 5
    position = (tick Mod (barWidth - blockWidth + 1)) + 1
    progressBar = "[" & String$(position - 1, ChrW(183)) & String$(blockWidth, ChrW(9632)) & _
                  String$(barWidth - blockWidth - position + 1, ChrW(183)) & "]"
    elapsedSeconds = DateDiff("s", startedAt, Now)

    Application.StatusBar = "NMDC Document Index - " & phaseText & " " & progressBar & _
                            "  Elapsed: " & CStr(elapsedSeconds) & " s"
End Sub

Public Function NMDC_ConfigValue(ByVal keyName As String) As String
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim row As ListRow
    Dim settingColumn As Long
    Dim valueColumn As Long

    Set ws = ThisWorkbook.Worksheets("Configuration")
    Set table = ws.ListObjects("Configuration")
    settingColumn = table.ListColumns("Setting").Index
    valueColumn = table.ListColumns("Current value").Index

    For Each row In table.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, settingColumn).Value)), keyName, vbTextCompare) = 0 Then
            NMDC_ConfigValue = CStr(row.Range.Cells(1, valueColumn).Value)
            Exit Function
        End If
    Next row
    NMDC_ConfigValue = ""
    Exit Function
Handler:
    NMDC_ConfigValue = ""
End Function

Public Sub NMDC_SetConfigValue(ByVal keyName As String, ByVal value As String)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim row As ListRow
    Dim settingColumn As Long
    Dim valueColumn As Long

    Set ws = ThisWorkbook.Worksheets("Configuration")
    Set table = ws.ListObjects("Configuration")
    settingColumn = table.ListColumns("Setting").Index
    valueColumn = table.ListColumns("Current value").Index

    For Each row In table.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, settingColumn).Value)), keyName, vbTextCompare) = 0 Then
            row.Range.Cells(1, valueColumn).Value = value
            Exit Sub
        End If
    Next row

    Set row = NMDC_BlankOrNewTableRow(table)
    row.Range.Cells(1, settingColumn).Value = keyName
    row.Range.Cells(1, valueColumn).Value = value
    Exit Sub
Handler:
    NMDC_LogError "CONFIG_WRITE_ERROR", "Excel could not save the configuration value.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LogError(ByVal code As String, ByVal friendlyMessage As String, ByVal technicalDetail As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim table As ListObject
    Dim row As ListRow

    Set ws = ThisWorkbook.Worksheets("Error Log")
    Set table = ws.ListObjects("ErrorLog")
    Set row = NMDC_BlankOrNewTableRow(table)

    NMDC_SetTableValue table, row, "Date/Time", Now
    NMDC_SetTableValue table, row, "Severity", "ERROR"
    NMDC_SetTableValue table, row, "Action", "EXCEL:" & code
    NMDC_SetTableValue table, row, "Plain-English Error", friendlyMessage
    NMDC_SetTableValue table, row, "Recommended Action", "Review the message and use Report Requirement / Problem if support is needed."
    NMDC_SetTableValue table, row, "Technical Detail", technicalDetail
    NMDC_SetTableValue table, row, "Run ID", ""
    NMDC_SetTableValue table, row, "Source File", ""
    NMDC_SetTableValue table, row, "Worksheet", ""
    NMDC_SetTableValue table, row, "Source Row/Cell", ""
    table.ListColumns("Date/Time").DataBodyRange.NumberFormat = "dd-mmm-yyyy hh:mm"
End Sub

Private Function NMDC_BlankOrNewTableRow(ByVal table As ListObject) As ListRow
    If table.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(table.ListRows(1).Range) = 0 Then
            Set NMDC_BlankOrNewTableRow = table.ListRows(1)
            Exit Function
        End If
    End If
    Set NMDC_BlankOrNewTableRow = table.ListRows.Add
End Function

Private Sub NMDC_SetTableValue(ByVal table As ListObject, ByVal row As ListRow, ByVal headerName As String, ByVal value As Variant)
    Dim columnIndex As Long
    columnIndex = table.ListColumns(headerName).Index
    row.Range.Cells(1, columnIndex).Value = value
End Sub
