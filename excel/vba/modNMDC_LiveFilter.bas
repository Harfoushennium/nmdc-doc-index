Attribute VB_Name = "modNMDC_LiveFilter"
Option Explicit

Private Const NMDC_SEARCH_BOX As String = "TxtBox_Search"
Private Const NMDC_SELECT_BUTTON As String = "NMDC_LiveFilter_Select"
Private Const NMDC_RESET_BUTTON As String = "NMDC_LiveFilter_Reset"
Private Const NMDC_ANCHOR_PREFIX As String = "_NMDC_LiveFilter_"
Private mIgnoreSearchEvent As Boolean

' Owner-approved dynamic search design.
' This intentionally follows the supplied Dynamic Live Filter forensic module:
'   * an ActiveX TextBox fires on every keystroke;
'   * Ctrl+Shift+F or the Select Column button chooses ONE table column;
'   * normal text performs a partial, case-insensitive native AutoFilter;
'   * text wrapped in double quotes performs exact, case-sensitive matching;
'   * Reset clears the live search and table filters.
' There is no all-column row concatenation and no hidden helper column.

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
        If Not ws Is Nothing Then NMDC_LiveFilterInstallSheet ws
    Next item

    NMDC_LiveFilterWake
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_INIT_ERROR", _
        "Excel could not initialize the dynamic Live Filter controls.", _
        Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "NMDC Live Filter", Err.Description
End Sub

Public Sub NMDC_LiveFilterWake()
    On Error Resume Next
    Application.OnKey "^+F", "NMDC_LiveFilterChooseColumn"
    If Not ActiveSheet Is Nothing Then
        If TypeName(ActiveSheet) = "Worksheet" Then NMDC_LiveFilterRefreshTargetCaption ActiveSheet
    End If
    On Error GoTo 0
End Sub

Public Sub NMDC_ApplyLiveFilterUX()
    ' Backward-compatible entry point retained for startup/owner UX callers.
    NMDC_LiveFilterWake
End Sub

Private Sub NMDC_LiveFilterInstallSheet(ByVal ws As Worksheet)
    On Error GoTo Handler

    Dim table As ListObject
    Dim area As Range
    Dim box As OLEObject
    Dim button As Shape

    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Range("A3:J3").UnMerge
    On Error GoTo Handler
    ws.Range("A3:J3").ClearContents
    ws.Rows(3).RowHeight = 26

    With ws.Range("A3")
        .Value = "LIVE FILTER"
        .Font.Name = "Aptos"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(31, 70, 90)
        .VerticalAlignment = xlCenter
    End With

    Set area = ws.Range("B3:E3")
    Set box = Nothing
    On Error Resume Next
    Set box = ws.OLEObjects(NMDC_SEARCH_BOX)
    On Error GoTo Handler

    If box Is Nothing Then
        Set box = ws.OLEObjects.Add( _
            ClassType:="Forms.TextBox.1", _
            Link:=False, _
            DisplayAsIcon:=False, _
            Left:=area.Left + 1, _
            Top:=area.Top + 1, _
            Width:=area.Width - 2, _
            Height:=area.Height - 2)
        box.Name = NMDC_SEARCH_BOX
    Else
        box.Left = area.Left + 1
        box.Top = area.Top + 1
        box.Width = area.Width - 2
        box.Height = area.Height - 2
    End If

    box.Placement = xlMoveAndSize
    box.PrintObject = False
    With box.Object
        .Font.Name = "Aptos"
        .Font.Size = 10
        .ForeColor = RGB(0, 0, 0)
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
        .MultiLine = False
        .EnterKeyBehavior = False
        .ControlTipText = "Type to filter the selected column instantly. Use Select Column or Ctrl+Shift+F first."
    End With

    On Error Resume Next
    ws.Shapes(NMDC_SELECT_BUTTON).Delete
    ws.Shapes(NMDC_RESET_BUTTON).Delete
    On Error GoTo Handler

    Set area = ws.Range("F3:H3")
    Set button = ws.Shapes.AddShape(5, area.Left + 1, area.Top + 1, area.Width - 2, area.Height - 2)
    button.Name = NMDC_SELECT_BUTTON
    button.OnAction = "NMDC_LiveFilterChooseColumn"
    NMDC_FormatLiveFilterButton button, RGB(20, 108, 148)

    Set area = ws.Range("I3:J3")
    Set button = ws.Shapes.AddShape(5, area.Left + 1, area.Top + 1, area.Width - 2, area.Height - 2)
    button.Name = NMDC_RESET_BUTTON
    button.OnAction = "NMDC_LiveFilterClear"
    button.TextFrame.Characters.Text = "RESET"
    NMDC_FormatLiveFilterButton button, RGB(91, 100, 112)

    NMDC_LiveFilterInstallChangeEvent ws
    NMDC_LiveFilterRefreshTargetCaption ws
    Exit Sub

Handler:
    Err.Raise Err.Number, "NMDC Live Filter - " & ws.Name, Err.Description
End Sub

Private Sub NMDC_FormatLiveFilterButton(ByVal button As Shape, ByVal fillColor As Long)
    With button
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = 3
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = fillColor
        .TextFrame.Characters.Font.Name = "Aptos"
        .TextFrame.Characters.Font.Size = 8.5
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Placement = xlMoveAndSize
    End With
End Sub

Private Sub NMDC_LiveFilterInstallChangeEvent(ByVal ws As Worksheet)
    On Error GoTo Handler

    Dim component As Object
    Dim codeModule As Object
    Dim sourceText As String
    Dim eventCode As String

    Set component = ThisWorkbook.VBProject.VBComponents(ws.CodeName)
    Set codeModule = component.CodeModule
    sourceText = ""
    If codeModule.CountOfLines > 0 Then
        sourceText = codeModule.Lines(1, codeModule.CountOfLines)
    End If

    If InStr(1, sourceText, "Private Sub TxtBox_Search_Change", vbTextCompare) > 0 Then Exit Sub

    eventCode = vbCrLf & _
        "Private Sub TxtBox_Search_Change()" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    NMDC_LiveFilterTextChanged Me" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "End Sub" & vbCrLf
    codeModule.AddFromString eventCode
    Exit Sub

Handler:
    Err.Raise Err.Number, "NMDC Live Filter event installer", Err.Description
End Sub

Public Sub NMDC_LiveFilterTextChanged(ByVal ws As Worksheet)
    On Error GoTo Handler
    If mIgnoreSearchEvent Then Exit Sub
    NMDC_LiveFilterApplyForSheet ws
    Exit Sub
Handler:
    NMDC_LogError "LIVE_FILTER_CHANGE_ERROR", _
        "Excel could not update the Live Filter while you typed on " & ws.Name & ".", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterChooseColumn()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim picked As Range
    Dim hit As Range
    Dim targetColumn As ListColumn
    Dim oldTarget As Range

    If ActiveSheet Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    Set table = NMDC_LiveFilterTableForSheet(ws)

    If table Is Nothing Then
        MsgBox "Live Filter is available on the main NMDC data/review tables only.", _
               vbInformation, "NMDC Document Index"
        Exit Sub
    End If

    On Error Resume Next
    Set picked = Application.InputBox( _
        "Click the HEADER of the column you want to search.", _
        "Select Live Filter Column", Type:=8)
    On Error GoTo Handler
    If picked Is Nothing Then Exit Sub

    Set hit = Intersect(picked.Cells(1, 1), table.HeaderRowRange)
    If hit Is Nothing Then
        MsgBox "Please click a header cell inside the main table on this sheet.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    Set targetColumn = table.ListColumns(CStr(hit.Value))
    If targetColumn Is Nothing Then Exit Sub
    If Left$(CStr(targetColumn.Name), 7) = "__NMDC_" Then
        MsgBox "Please choose a normal user-visible column.", vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    Set oldTarget = NMDC_LiveFilterTargetRange(ws)
    If Not oldTarget Is Nothing Then NMDC_LiveFilterClearTargetFilter oldTarget

    NMDC_LiveFilterSetTarget ws, targetColumn
    NMDC_LiveFilterRefreshTargetCaption ws
    NMDC_LiveFilterApplyForSheet ws

    On Error Resume Next
    ws.OLEObjects(NMDC_SEARCH_BOX).Activate
    On Error GoTo 0
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_TARGET_ERROR", _
        "Excel could not change the Live Filter target column.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterClear()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim box As OLEObject

    If ActiveSheet Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Exit Sub

    mIgnoreSearchEvent = True
    Set box = Nothing
    On Error Resume Next
    Set box = ws.OLEObjects(NMDC_SEARCH_BOX)
    If Not box Is Nothing Then box.Object.Value = ""
    If table.AutoFilter.FilterMode Then table.AutoFilter.ShowAllData
    On Error GoTo Handler
    mIgnoreSearchEvent = False

    If Not box Is Nothing Then box.Activate
    Exit Sub

Handler:
    mIgnoreSearchEvent = False
    NMDC_LogError "LIVE_FILTER_CLEAR_ERROR", _
        "Excel could not reset the Live Filter.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterApplyForSheet(ByVal ws As Worksheet)
    On Error GoTo Handler

    Dim targetRange As Range
    Dim table As ListObject
    Dim box As OLEObject
    Dim userText As String
    Dim cleanText As String
    Dim fieldNumber As Long
    Dim strictMode As Boolean
    Dim visibleRows As Range
    Dim cell As Range
    Dim previousScreen As Boolean

    Set targetRange = NMDC_LiveFilterTargetRange(ws)
    If targetRange Is Nothing Then
        Application.StatusBar = "NMDC Live Filter: select a search column first (button or Ctrl+Shift+F)."
        Exit Sub
    End If

    Set table = targetRange.ListObject
    If table Is Nothing Then Exit Sub

    Set box = Nothing
    On Error Resume Next
    Set box = ws.OLEObjects(NMDC_SEARCH_BOX)
    On Error GoTo Handler
    If box Is Nothing Then Exit Sub

    userText = CStr(box.Object.Value)
    cleanText = userText
    strictMode = False
    If Len(cleanText) > 1 Then
        If Left$(cleanText, 1) = Chr$(34) And Right$(cleanText, 1) = Chr$(34) Then
            strictMode = True
            cleanText = Mid$(cleanText, 2, Len(cleanText) - 2)
        End If
    End If

    fieldNumber = targetRange.Column - table.Range.Column + 1
    If fieldNumber < 1 Or fieldNumber > table.ListColumns.Count Then Exit Sub

    previousScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' Clear only the previous Live Filter criterion for the selected field.
    On Error Resume Next
    table.Range.AutoFilter Field:=fieldNumber
    table.DataBodyRange.EntireRow.Hidden = False
    On Error GoTo Handler

    If Len(Trim$(cleanText)) = 0 Then GoTo CleanExit

    If strictMode Then
        table.Range.AutoFilter Field:=fieldNumber, Criteria1:=cleanText
        Set visibleRows = Nothing
        On Error Resume Next
        Set visibleRows = targetRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo Handler
        If Not visibleRows Is Nothing Then
            For Each cell In visibleRows.Cells
                If StrComp(CStr(cell.Value), cleanText, vbBinaryCompare) <> 0 Then
                    cell.EntireRow.Hidden = True
                End If
            Next cell
        End If
    Else
        table.Range.AutoFilter Field:=fieldNumber, _
            Criteria1:="=*" & NMDC_LiveFilterEscapeWildcards(cleanText) & "*"
    End If

CleanExit:
    Application.ScreenUpdating = previousScreen
    Application.StatusBar = False
    Exit Sub

Handler:
    Application.ScreenUpdating = previousScreen
    Application.StatusBar = False
    NMDC_LogError "LIVE_FILTER_APPLY_ERROR", _
        "Excel could not apply the Live Filter on " & ws.Name & ".", _
        Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_LiveFilterEscapeWildcards(ByVal value As String) As String
    value = Replace(value, "~", "~~")
    value = Replace(value, "*", "~*")
    value = Replace(value, "?", "~?")
    NMDC_LiveFilterEscapeWildcards = value
End Function

Private Sub NMDC_LiveFilterSetTarget(ByVal ws As Worksheet, ByVal targetColumn As ListColumn)
    Dim anchorName As String
    anchorName = NMDC_LiveFilterAnchorName(ws)

    On Error Resume Next
    ThisWorkbook.Names(anchorName).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:=anchorName, _
        RefersTo:=targetColumn.DataBodyRange, _
        Visible:=False
End Sub

Private Function NMDC_LiveFilterTargetRange(ByVal ws As Worksheet) As Range
    Dim anchorName As String
    anchorName = NMDC_LiveFilterAnchorName(ws)

    On Error Resume Next
    Set NMDC_LiveFilterTargetRange = ThisWorkbook.Names(anchorName).RefersToRange
    On Error GoTo 0
End Function

Private Function NMDC_LiveFilterAnchorName(ByVal ws As Worksheet) As String
    NMDC_LiveFilterAnchorName = NMDC_ANCHOR_PREFIX & ws.CodeName
End Function

Private Sub NMDC_LiveFilterClearTargetFilter(ByVal targetRange As Range)
    On Error Resume Next
    Dim table As ListObject
    Dim fieldNumber As Long
    Set table = targetRange.ListObject
    If table Is Nothing Then Exit Sub
    fieldNumber = targetRange.Column - table.Range.Column + 1
    table.Range.AutoFilter Field:=fieldNumber
    table.DataBodyRange.EntireRow.Hidden = False
    On Error GoTo 0
End Sub

Private Sub NMDC_LiveFilterRefreshTargetCaption(ByVal ws As Worksheet)
    On Error Resume Next

    Dim button As Shape
    Dim targetRange As Range
    Dim targetName As String

    Set button = ws.Shapes(NMDC_SELECT_BUTTON)
    If button Is Nothing Then Exit Sub

    Set targetRange = NMDC_LiveFilterTargetRange(ws)
    If targetRange Is Nothing Then
        button.TextFrame.Characters.Text = "SELECT COLUMN"
        button.Fill.ForeColor.RGB = RGB(194, 139, 0)
        button.Line.ForeColor.RGB = RGB(194, 139, 0)
    Else
        targetName = CStr(targetRange.ListObject.HeaderRowRange.Cells( _
            1, targetRange.Column - targetRange.ListObject.Range.Column + 1).Value)
        button.TextFrame.Characters.Text = "COLUMN: " & targetName
        button.Fill.ForeColor.RGB = RGB(20, 108, 148)
        button.Line.ForeColor.RGB = RGB(20, 108, 148)
    End If
    On Error GoTo 0
End Sub

Public Sub NMDC_LiveFilterSheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Compatibility hook retained for existing ThisWorkbook event wiring.
    ' Dynamic filtering is now driven by the ActiveX TextBox Change event.
End Sub

Public Sub NMDC_LiveFilterSheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If NMDC_LiveFilterTableForSheet(Sh) Is Nothing Then Exit Sub
    NMDC_LiveFilterRefreshTargetCaption Sh
    If StrComp(CStr(Sh.Name), "Pending Update", vbTextCompare) = 0 Then
        Application.Run "NMDC_RebuildPendingSourcePanel"
    End If
    On Error GoTo 0
End Sub

Private Function NMDC_LiveFilterTableForSheet(ByVal ws As Worksheet) As ListObject
    Dim tableName As String

    Select Case UCase$(Trim$(ws.Name))
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
