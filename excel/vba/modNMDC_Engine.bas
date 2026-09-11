Attribute VB_Name = "modNMDC_Engine"
Option Explicit

Private Const ENGINE_RELATIVE_PATH As String = "engine\nmdc_index_engine.exe"
Private Const RUNTIME_RELATIVE_PATH As String = "runtime"
Private Const EXCHANGE_RELATIVE_PATH As String = "runtime\excel_exchange"

Public Function NMDC_WorkbookFolder() As String
    NMDC_WorkbookFolder = ThisWorkbook.Path
End Function

Public Function NMDC_EnginePath() As String
    Dim configured As String
    configured = Trim$(NMDC_ConfigValue("Engine Executable Path"))
    If Len(configured) > 0 Then
        NMDC_EnginePath = configured
    Else
        NMDC_EnginePath = NMDC_WorkbookFolder() & "\" & ENGINE_RELATIVE_PATH
    End If
End Function

Public Function NMDC_RuntimePath() As String
    Dim configured As String
    configured = Trim$(NMDC_ConfigValue("Runtime Folder"))
    If Len(configured) > 0 Then
        NMDC_RuntimePath = configured
    Else
        NMDC_RuntimePath = NMDC_WorkbookFolder() & "\" & RUNTIME_RELATIVE_PATH
    End If
End Function

Public Function NMDC_ExchangePath() As String
    NMDC_ExchangePath = NMDC_WorkbookFolder() & "\" & EXCHANGE_RELATIVE_PATH
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

    If Len(Dir$(enginePath)) = 0 Then
        NMDC_LogError "ENGINE_MISSING", _
            "The NMDC Index engine could not be found.", _
            "Expected engine: " & enginePath & ". Contact support or use Report Requirement / Problem."
        MsgBox "The NMDC Index engine could not be found." & vbCrLf & vbCrLf & _
               "Nothing was changed. Please open Error Log or use Report Requirement / Problem.", _
               vbExclamation, "NMDC Document Index"
        NMDC_RunEngine = 9001
        Exit Function
    End If

    cmd = NMDC_Quote(enginePath) & " " & commandName & _
          " --state-dir " & NMDC_Quote(runtimePath) & _
          " --exchange-dir " & NMDC_Quote(exchangePath)
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
        Err.Number & " - " & Err.Description
    NMDC_RunEngine = 9002
End Function

Public Function NMDC_ConfigValue(ByVal keyName As String) As String
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim hit As Range

    Set ws = ThisWorkbook.Worksheets("Configuration")
    Set hit = ws.Columns("A").Find(What:=keyName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not hit Is Nothing Then
        NMDC_ConfigValue = CStr(hit.Offset(0, 1).Value)
    Else
        NMDC_ConfigValue = ""
    End If
    Exit Function
Handler:
    NMDC_ConfigValue = ""
End Function

Public Sub NMDC_SetConfigValue(ByVal keyName As String, ByVal value As String)
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim hit As Range
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("Configuration")
    Set hit = ws.Columns("A").Find(What:=keyName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If hit Is Nothing Then
        nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(nextRow, "A").Value = keyName
        ws.Cells(nextRow, "B").Value = value
    Else
        hit.Offset(0, 1).Value = value
    End If
    Exit Sub
Handler:
    NMDC_LogError "CONFIG_WRITE_ERROR", "Excel could not save the configuration value.", Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LogError(ByVal code As String, ByVal friendlyMessage As String, ByVal technicalDetail As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets("Error Log")
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = "ERROR"
    ws.Cells(nextRow, 3).Value = code
    ws.Cells(nextRow, 4).Value = friendlyMessage
    ws.Cells(nextRow, 5).Value = "Review the message and use Report Requirement / Problem if support is needed."
    ws.Cells(nextRow, 6).Value = technicalDetail
End Sub
