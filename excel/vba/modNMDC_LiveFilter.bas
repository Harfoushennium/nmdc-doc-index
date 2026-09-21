Attribute VB_Name = "Mod_LiveFilter"
Option Explicit

' ==============================================================================
' MODULE: Mod_LiveFilter
' DESCRIPTION: NMDC integration of the owner's Dynamic Live Filter Tool REV03.
' ==============================================================================

Public TheListener As Cls_LiveFilter_Listener

Private Const SEARCH_BOX_NAME As String = "TxtBox_Search"
Private Const RESET_BUTTON_NAME As String = "Btn_Reset_Search"

Private Sub Auto_Open()
    Call NMDC_LiveFilterInitialize
End Sub

Private Sub Auto_Close()
    On Error Resume Next
    Application.OnKey "^+F"
    On Error GoTo 0
End Sub

Public Sub NMDC_LiveFilterInitialize()
    On Error GoTo Handler

    Dim sheetNames As Variant
    Dim item As Variant
    Dim ws As Worksheet

    sheetNames = Array( _
        "Master Documents", "Revisions", "Transactions", "Pending Update", _
        "Review Flags", "User Decisions", "Update History", "Error Log")

    For Each item In sheetNames
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(item))
        On Error GoTo Handler
        If Not ws Is Nothing Then NMDC_EnsureReferenceLiveFilterControls ws
    Next item

    Call Enable_The_Shortcut_Safe

    If Not ActiveSheet Is Nothing Then
        If TypeName(ActiveSheet) = "Worksheet" Then
            If Not NMDC_LiveFilterTableForSheet(ActiveSheet) Is Nothing Then
                Call Wake_Up_The_Tool(False)
            End If
        End If
    End If
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_INIT_ERROR", _
        "Excel could not initialize the owner-reference Live Filter controls.", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_EnsureReferenceLiveFilterControls(ByVal ws As Worksheet)
    On Error GoTo Handler

    Dim TheOle As OLEObject
    Dim TheBox As Object
    Dim area As Range
    Dim resetButton As Button

    If NMDC_LiveFilterTableForSheet(ws) Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Shapes("NMDC_LiveFilter_Open").Delete
    ws.Shapes("NMDC_LiveFilter_Reset").Delete
    On Error GoTo Handler

    Set TheOle = Nothing
    On Error Resume Next
    Set TheOle = ws.OLEObjects(SEARCH_BOX_NAME)
    On Error GoTo Handler

    Set area = ws.Range("B3:J3")

    If TheOle Is Nothing Then
        Set TheOle = ws.OLEObjects.Add( _
            ClassType:="Forms.TextBox.1", _
            Link:=False, _
            DisplayAsIcon:=False, _
            Left:=area.Left + 1, _
            Top:=area.Top + 1, _
            Width:=area.Width - 2, _
            Height:=area.Height - 2)
        TheOle.Name = SEARCH_BOX_NAME
    Else
        TheOle.Left = area.Left + 1
        TheOle.Top = area.Top + 1
        TheOle.Width = area.Width - 2
        TheOle.Height = area.Height - 2
    End If

    TheOle.Placement = xlMoveAndSize
    Set TheBox = TheOle.Object
    On Error Resume Next
    TheBox.MultiLine = False
    TheBox.EnterKeyBehavior = False
    TheBox.WordWrap = False
    TheBox.BorderStyle = 1
    TheBox.SpecialEffect = 0
    TheBox.BackColor = RGB(255, 255, 255)
    TheBox.Font.Name = "Calibri"
    TheBox.Font.Size = 11
    On Error GoTo Handler

    Set resetButton = Nothing
    On Error Resume Next
    Set resetButton = ws.Buttons(RESET_BUTTON_NAME)
    On Error GoTo Handler

    Set area = ws.Range("K3:L3")
    If resetButton Is Nothing Then
        Set resetButton = ws.Buttons.Add(area.Left + 1, area.Top + 1, area.Width - 2, area.Height - 2)
        resetButton.Name = RESET_BUTTON_NAME
    Else
        resetButton.Left = area.Left + 1
        resetButton.Top = area.Top + 1
        resetButton.Width = area.Width - 2
        resetButton.Height = area.Height - 2
    End If

    resetButton.Caption = "RESET SEARCH"
    resetButton.OnAction = "Reset_Search_Click"
    resetButton.Placement = xlMoveAndSize
    On Error Resume Next
    resetButton.Font.Name = "Calibri"
    resetButton.Font.Size = 10
    resetButton.Font.Bold = True
    On Error GoTo Handler

    ws.Rows(3).RowHeight = 31.5
    Exit Sub

Handler:
    Err.Raise Err.Number, "NMDC Live Filter control setup on " & ws.Name, Err.Description
End Sub

Public Sub Wake_Up_The_Tool(Optional ShowErrors As Boolean = False)
    Dim ws As Worksheet
    Dim TheBox As Object

    If ActiveSheet Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    If NMDC_LiveFilterTableForSheet(ws) Is Nothing Then Exit Sub

    On Error Resume Next
    Set TheBox = ws.OLEObjects(SEARCH_BOX_NAME).Object
    If Err.Number <> 0 Or TheBox Is Nothing Then
        Err.Clear
        NMDC_EnsureReferenceLiveFilterControls ws
        Set TheBox = ws.OLEObjects(SEARCH_BOX_NAME).Object
    End If

    If Err.Number <> 0 Or TheBox Is Nothing Then
        On Error GoTo 0
        If ShowErrors Then
            MsgBox "MISSING OBJECT: '" & SEARCH_BOX_NAME & "'" & vbCrLf & _
                   "The tool cannot find the search box. Please run setup again.", vbCritical
        End If
        Exit Sub
    End If
    On Error GoTo 0

    Set TheListener = New Cls_LiveFilter_Listener
    Set TheListener.ExcelApp = Application
    Set TheListener.SearchBox = TheBox

    Call Enable_The_Shortcut_Safe
    Call Fix_Placeholder_Text(TheBox)
End Sub

Public Sub Reset_Search_Click()
    Dim SavedLocation As Range
    Dim TargetTable As ListObject

    Call Wake_Up_The_Tool(ShowErrors:=True)
    If TheListener Is Nothing Then Exit Sub

    On Error Resume Next
    Set SavedLocation = ThisWorkbook.Names("LiveFilter_Anchor").RefersToRange
    If Err.Number <> 0 Or SavedLocation Is Nothing Then
        Err.Clear
        On Error GoTo 0
        MsgBox "CONFIGURATION MISSING" & vbCrLf & _
               "Press 'Ctrl + Shift + F' to select a column first.", vbExclamation
        Exit Sub
    End If

    Set TargetTable = SavedLocation.ListObject
    If TargetTable Is Nothing Then
        On Error GoTo 0
        MsgBox "BROKEN TABLE LINK" & vbCrLf & "Target column is not inside a table.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    If TargetTable.ShowAutoFilter Then
        If TargetTable.AutoFilter.FilterMode Then TargetTable.AutoFilter.ShowAllData
    End If

    If Not TheListener.SearchBox Is Nothing Then
        TheListener.SearchBox.Value = ""
        Call Fix_Placeholder_Text(TheListener.SearchBox)
        On Error Resume Next
        TheListener.SearchBox.Activate
        On Error GoTo 0
    End If
End Sub

Private Sub Enable_The_Shortcut_Safe()
    On Error Resume Next
    Application.OnKey "^+F", "Ask_User_For_Target_Column"
    On Error GoTo 0
End Sub

Public Sub Ask_User_For_Target_Column()
    Dim UserSelection As Range
    On Error Resume Next
    Set UserSelection = Application.InputBox("Click the HEADER of the column to filter:", "Set Search Target", Type:=8)
    On Error GoTo 0
    If UserSelection Is Nothing Then Exit Sub
    Call Setup_New_Target(UserSelection)
End Sub

Public Sub Setup_New_Target(TheHeader As Range)
    Dim TheTable As ListObject
    Dim TheColumn As ListColumn

    On Error Resume Next
    Set TheTable = TheHeader.ListObject
    If TheTable Is Nothing Then
        On Error GoTo 0
        MsgBox "Click inside an EXCEL TABLE.", vbExclamation
        Exit Sub
    End If

    If Intersect(TheHeader, TheTable.HeaderRowRange) Is Nothing Then
        On Error GoTo 0
        MsgBox "Click the HEADER row.", vbExclamation
        Exit Sub
    End If

    Set TheColumn = TheTable.ListColumns(TheHeader.Value)
    If TheColumn Is Nothing Then
        On Error GoTo 0
        MsgBox "Column not found.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    If Left$(CStr(TheColumn.Name), 7) = "__NMDC_" Then
        MsgBox "Choose a normal user-visible column.", vbExclamation
        Exit Sub
    End If

    If TheColumn.DataBodyRange Is Nothing Then
        MsgBox "This table does not contain data rows yet.", vbInformation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Names("LiveFilter_Anchor").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="LiveFilter_Anchor", RefersTo:=TheColumn.DataBodyRange, Visible:=False

    TheTable.Parent.Activate
    Call Wake_Up_The_Tool(ShowErrors:=True)
    If Not TheListener Is Nothing Then
        Call Fix_Placeholder_Text(TheListener.SearchBox, CStr(TheHeader.Value))
    End If
End Sub

Public Sub Run_Live_Filter(ByVal UserText As String)
    Dim TheTable As ListObject, TargetRange As Range
    Dim VisibleRows As Range, Cell As Range, RowsToHide As Range
    Dim ColumnNumber As Long
    Dim CleanText As String
    Dim rawTokens() As String
    Dim posTerms() As String, negTerms() As String
    Dim pCount As Long, nCount As Long
    Dim token As String, tempStr As String
    Dim i As Long
    Dim MatchFailed As Boolean

    If InStr(1, UserText, "(Ctrl+Shift+F)") > 0 Then Exit Sub

    On Error Resume Next
    Set TargetRange = ThisWorkbook.Names("LiveFilter_Anchor").RefersToRange
    If TargetRange Is Nothing Then Exit Sub
    Set TheTable = TargetRange.ListObject
    If TheTable Is Nothing Then Exit Sub
    On Error GoTo 0

    ColumnNumber = TargetRange.Column - TheTable.DataBodyRange.Column + 1
    Application.ScreenUpdating = False

    If Left(UserText, 1) = """" And Right(UserText, 1) = """" And Len(UserText) > 1 Then
        CleanText = Mid(UserText, 2, Len(UserText) - 2)
        TheTable.Range.AutoFilter Field:=ColumnNumber, Criteria1:=CleanText

        On Error Resume Next
        Set VisibleRows = TargetRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not VisibleRows Is Nothing Then
            For Each Cell In VisibleRows
                If StrComp(Cell.Value, CleanText, vbBinaryCompare) <> 0 Then
                    If RowsToHide Is Nothing Then Set RowsToHide = Cell Else Set RowsToHide = Union(RowsToHide, Cell)
                End If
            Next Cell
            If Not RowsToHide Is Nothing Then RowsToHide.EntireRow.Hidden = True
        End If
        Application.ScreenUpdating = True
        Exit Sub
    End If

    CleanText = Trim(UserText)

    If CleanText = "" Then
        TheTable.Range.AutoFilter Field:=ColumnNumber
        Application.ScreenUpdating = True
        Exit Sub
    End If

    tempStr = Replace(CleanText, "+", " ")
    rawTokens = Split(tempStr, " ")

    pCount = 0: nCount = 0
    ReDim posTerms(0 To UBound(rawTokens))
    ReDim negTerms(0 To UBound(rawTokens))

    For i = LBound(rawTokens) To UBound(rawTokens)
        token = Trim(rawTokens(i))
        If Len(token) > 0 Then
            If Left(token, 1) = "-" And Len(token) > 1 Then
                negTerms(nCount) = Mid(token, 2)
                nCount = nCount + 1
            Else
                posTerms(pCount) = token
                pCount = pCount + 1
            End If
        End If
    Next i

    If pCount > 0 Then
        TheTable.Range.AutoFilter Field:=ColumnNumber, Criteria1:="=*" & posTerms(0) & "*"
    Else
        TheTable.Range.AutoFilter Field:=ColumnNumber
    End If

    If pCount > 1 Or nCount > 0 Then
        On Error Resume Next
        Set VisibleRows = TargetRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not VisibleRows Is Nothing Then
            For Each Cell In VisibleRows
                MatchFailed = False

                If pCount > 1 Then
                    For i = 1 To pCount - 1
                        If InStr(1, Cell.Value, posTerms(i), vbTextCompare) = 0 Then
                            MatchFailed = True
                            Exit For
                        End If
                    Next i
                End If

                If Not MatchFailed And nCount > 0 Then
                    For i = 0 To nCount - 1
                        If InStr(1, Cell.Value, negTerms(i), vbTextCompare) > 0 Then
                            MatchFailed = True
                            Exit For
                        End If
                    Next i
                End If

                If MatchFailed Then
                    If RowsToHide Is Nothing Then Set RowsToHide = Cell Else Set RowsToHide = Union(RowsToHide, Cell)
                End If
            Next Cell

            If Not RowsToHide Is Nothing Then RowsToHide.EntireRow.Hidden = True
        End If
    End If

    Application.ScreenUpdating = True
End Sub

Public Sub Fix_Placeholder_Text(TheBox As Object, Optional ColName As String = "")
    If TheBox Is Nothing Then Exit Sub

    If ColName = "" Then
        On Error Resume Next
        ColName = ThisWorkbook.Names("LiveFilter_Anchor").RefersToRange.ListObject.HeaderRowRange.Cells(1, _
                  ThisWorkbook.Names("LiveFilter_Anchor").RefersToRange.Column - _
                  ThisWorkbook.Names("LiveFilter_Anchor").RefersToRange.ListObject.DataBodyRange.Column + 1).Value
        On Error GoTo 0
    End If
    If ColName = "" Then ColName = "Target"

    On Error Resume Next
    If Trim(TheBox.Value) = "" Or TheBox.Value Like "Search *" Then
        Application.ScreenUpdating = False
        TheBox.ForeColor = RGB(150, 150, 150)
        TheBox.Value = "Search " & ColName & "... (+AND / -EXCLUDE) (Ctrl+Shift+F: Change Column)"
        Application.ScreenUpdating = True
    End If
    On Error GoTo 0
End Sub

Public Sub Clear_Placeholder_Text(TheBox As Object)
    If TheBox Is Nothing Then Exit Sub
    On Error Resume Next
    If TheBox.ForeColor = RGB(150, 150, 150) Then
        TheBox.Value = ""
        TheBox.ForeColor = RGB(0, 0, 0)
    End If
    On Error GoTo 0
End Sub

Public Sub NMDC_LiveFilterWake()
    Call Wake_Up_The_Tool(False)
End Sub

Public Sub NMDC_LiveFilterSheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If NMDC_LiveFilterTableForSheet(Sh) Is Nothing Then Exit Sub
    NMDC_EnsureReferenceLiveFilterControls Sh
    Sh.Activate
    Call Wake_Up_The_Tool(False)
    On Error GoTo 0
End Sub

Public Sub NMDC_LiveFilterSheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Native ActiveX SearchBox_Change is the typing event.
End Sub

Public Sub NMDC_LiveFilterSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ' Native ActiveX textbox receives the click/caret directly.
End Sub

Public Sub NMDC_LiveFilterChooseColumn()
    Ask_User_For_Target_Column
End Sub

Public Sub NMDC_LiveFilterClear()
    Reset_Search_Click
End Sub

Public Sub NMDC_LiveFilterShow()
    On Error Resume Next
    Wake_Up_The_Tool True
    If Not TheListener Is Nothing Then TheListener.SearchBox.Activate
    On Error GoTo 0
End Sub

Public Sub NMDC_LiveFilterFormChanged(ByVal sheetName As String, ByVal queryText As String)
    If StrComp(ActiveSheet.Name, sheetName, vbTextCompare) = 0 Then Run_Live_Filter queryText
End Sub

Public Sub NMDC_LiveFilterSelectColumnByName(ByVal sheetName As String, ByVal columnName As String)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim targetColumn As ListColumn

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Err.Raise vbObjectError + 951, "NMDC Live Filter", "No searchable table on " & sheetName

    Set targetColumn = table.ListColumns(columnName)
    If targetColumn Is Nothing Then Err.Raise vbObjectError + 952, "NMDC Live Filter", "Unknown search column: " & columnName
    If Left$(CStr(targetColumn.Name), 7) = "__NMDC_" Then Err.Raise vbObjectError + 953, "NMDC Live Filter", "Internal helper columns cannot be searched."

    ws.Activate
    Setup_New_Target targetColumn.Range.Cells(1, 1)
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_TARGET_NAME_ERROR", _
        "Excel could not set the Live Filter target column.", _
        Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "NMDC Live Filter", Err.Description
End Sub

Private Function NMDC_LiveFilterTableForSheet(ByVal ws As Worksheet) As ListObject
    Dim tableName As String

    Select Case UCase$(ws.Name)
        Case "MASTER DOCUMENTS": tableName = "MasterDocuments"
        Case "REVISIONS": tableName = "RevisionRegister"
        Case "TRANSACTIONS": tableName = "EventRegister"
        Case "PENDING UPDATE": tableName = "PendingUpdate"
        Case "REVIEW FLAGS": tableName = "ReviewFlags"
        Case "USER DECISIONS": tableName = "UserDecisionLog"
        Case "UPDATE HISTORY": tableName = "UpdateHistory"
        Case "ERROR LOG": tableName = "ErrorLog"
        Case Else: Exit Function
    End Select

    On Error Resume Next
    Set NMDC_LiveFilterTableForSheet = ws.ListObjects(tableName)
    On Error GoTo 0
End Function
