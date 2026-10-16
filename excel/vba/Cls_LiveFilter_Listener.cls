VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_LiveFilter_Listener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' CLASS: Cls_LiveFilter_Listener
' DESCRIPTION: Event Listener matching the owner's Dynamic Live Filter REV03.
' ==============================================================================

Public WithEvents ExcelApp As Application
Attribute ExcelApp.VB_VarHelpID = -1
Public WithEvents SearchBox As MSForms.TextBox
Attribute SearchBox.VB_VarHelpID = -1

Private Sub ExcelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If SearchBox Is Nothing Then Exit Sub
    If Trim(SearchBox.Value) = "" Then
        Call Mod_LiveFilter.Fix_Placeholder_Text(SearchBox)
    End If
    On Error GoTo 0
End Sub

Private Sub SearchBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    Call Mod_LiveFilter.Clear_Placeholder_Text(SearchBox)
    On Error GoTo 0
End Sub

Private Sub SearchBox_Change()
    On Error Resume Next
    If SearchBox.ForeColor = RGB(0, 0, 0) Then
        Call Mod_LiveFilter.Run_Live_Filter(SearchBox.Value)
    End If
    On Error GoTo 0
End Sub

Private Sub SearchBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    Call Mod_LiveFilter.Clear_Placeholder_Text(SearchBox)
    On Error GoTo 0
End Sub
