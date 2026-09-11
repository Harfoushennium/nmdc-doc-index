Attribute VB_Name = "modNMDC_Startup"
Option Explicit

Public Sub Auto_Open()
    On Error GoTo Handler

    Application.StatusBar = "NMDC Document Index: refreshing dashboard..."
    If NMDC_RunEngine("export-excel") = 0 Then
        NMDC_RefreshExchangeData
    End If
    Application.StatusBar = False
    Exit Sub

Handler:
    Application.StatusBar = False
    NMDC_LogError "STARTUP_REFRESH_ERROR", _
        "The workbook opened, but the NMDC Index dashboard could not refresh automatically.", _
        Err.Number & " - " & Err.Description
End Sub
