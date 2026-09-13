Attribute VB_Name = "modNMDC_Startup"
Option Explicit

Public Sub Auto_Open()
    On Error GoTo Handler

    Application.StatusBar = "NMDC Document Index: refreshing dashboard..."
    If NMDC_RunEngine("export-excel") = 0 Then
        If Not NMDC_RefreshExchangeData() Then
            NMDC_LogError "STARTUP_EXCHANGE_ERROR", _
                "The workbook opened, but one or more displayed tables could not be refreshed.", _
                "Open Error Log and press Refresh Dashboard before reviewing or approving an update."
        End If
    End If
    Application.StatusBar = False
    Exit Sub

Handler:
    Application.StatusBar = False
    NMDC_LogError "STARTUP_REFRESH_ERROR", _
        "The workbook opened, but the NMDC Index dashboard could not refresh automatically.", _
        Err.Number & " - " & Err.Description
End Sub
