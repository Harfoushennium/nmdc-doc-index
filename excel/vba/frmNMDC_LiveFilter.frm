VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNMDC_LiveFilter 
   Caption         =   "NMDC Live Filter"
   ClientHeight    =   2475
   ClientLeft      =   99
   ClientTop       =   429
   ClientWidth     =   7007
   OleObjectBlob   =   "frmNMDC_LiveFilter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNMDC_LiveFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSheetName As String

Public Sub NMDC_Bind(ByVal sheetName As String, ByVal currentText As String, ByVal targetName As String)
    mSheetName = sheetName
    Me.cmbSearch.Text = currentText
    Me.lblTarget.Caption = "Column: " & targetName
    Me.lblHelp.Caption = "Partial terms; spaces/+ = AND; -word = exclude; quotes = exact; * and ? = wildcard. Results update as you type."
End Sub

Public Sub NMDC_FocusSearch()
    On Error Resume Next
    Me.cmbSearch.SetFocus
    Me.cmbSearch.SelStart = Len(Me.cmbSearch.Text)
    On Error GoTo 0
End Sub

Public Sub NMDC_SetSearchText(ByVal value As String)
    Me.cmbSearch.Text = value
    NMDC_FocusSearch
End Sub

Private Sub UserForm_Activate()
    Me.cmbSearch.SetFocus
    Me.cmbSearch.SelStart = Len(Me.cmbSearch.Text)
End Sub

Private Sub cmbSearch_Change()
    If Len(mSheetName) = 0 Then Exit Sub
    NMDC_LiveFilterFormChanged mSheetName, Me.cmbSearch.Text
End Sub

Private Sub cmdChooseColumn_Click()
    NMDC_LiveFilterChooseColumn
End Sub

Private Sub cmdReset_Click()
    Me.cmbSearch.Text = vbNullString
    NMDC_LiveFilterClear
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

