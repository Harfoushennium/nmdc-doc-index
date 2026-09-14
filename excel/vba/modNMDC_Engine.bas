Attribute VB_Name = "modNMDC_Engine"
Option Explicit

Private Const ENGINE_RELATIVE_PATH As String = "engine\nmdc_index_engine.exe"
Private Const RUNTIME_RELATIVE_PATH As String = "runtime"
Private Const EXCHANGE_RELATIVE_PATH As String = "excel_exchange"
Private Const CONFIG_RELATIVE_PATH As String = "config"

Public Function NMDC_WorkbookFolder() As String
    NMDC_WorkbookFolder = ThisWorkbook.Path
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
    Dim exitCode As Long

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

    cmd = NMDC_Quote(enginePath) & " " & commandName & _
          " --state-dir " & NMDC_Quote(runtimePath) & _
          " --exchange-dir " & NMDC_Quote(exchangePath) & _
          " --config-dir " & NMDC_Quote(NMDC_ConfigPath())
    If Len(Trim$(extraArgs)) > 0 Then cmd = cmd & " " & extraArgs

    Set shell = CreateObject("WScript.Shell")
    ' WindowStyle=0 hides the console. WaitOnReturn=True keeps Excel informed of success/failure.
    exitCode = shell.Run(cmd, 0, True)
    NMDC_RunEngine = exitCode

    If exitCode <> 0 Then
        NMDC_LogError "ENGINE_EXIT_CODE", _
            "The requested action did not complete successfully.", _
            "Engine exit code: " & CStr(exitCode) & ". Command: " & commandName
    End If
    Exit Function

Handler:
    NMDC_LogError "VBA_ENGINE_LAUNCH_ERROR", _
        "Excel could not start the NMDC Index engine.", _
        Err.Number & " - " & Err.Description & "; " & NMDC_PathDiagnostics()
    NMDC_RunEngine = 9002
End Function

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
