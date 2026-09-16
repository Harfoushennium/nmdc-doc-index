Attribute VB_Name = "modNMDC_LiveFilter"
Option Explicit

Private Const NMDC_SELECT_BUTTON As String = "NMDC_LiveFilter_Select"
Private Const NMDC_RESET_BUTTON As String = "NMDC_LiveFilter_Reset"
Private Const NMDC_SEARCH_DISPLAY As String = "NMDC_LiveFilter_Display"
Private Const NMDC_ANCHOR_PREFIX As String = "_NMDC_LiveFilter_"

Private mCaptureActive As Boolean
Private mCaptureSheetName As String
Private mQuery As String

' Reliable dynamic Live Filter for current Microsoft 365.
'
' IMPORTANT:
' Worksheet ActiveX text boxes are intentionally NOT used. Microsoft 365 /
' Office 2024 disable ActiveX controls by default and the owner environment
' already returned runtime error 40040 while creating Forms.TextBox.1.
'
' To keep true per-keystroke behaviour without ActiveX, this module uses
' Application.OnKey only while Live Filter capture mode is active:
'   1) choose one table column;
'   2) type normally - each key immediately updates the native table filter;
'   3) Backspace edits the query; Delete clears it; Esc/Enter stops capture;
'   4) click the search display to resume capture for the same column.
'
' The key hooks are temporary and are released as soon as capture stops.

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
        "Excel could not initialize the Live Filter interface.", _
        Err.Number & " - " & Err.Description
    ' Live Filter must never block workbook creation or scanning.
End Sub

Public Sub NMDC_LiveFilterWake()
    On Error Resume Next
    Application.OnKey "^+F", NMDC_LiveFilterQualifiedMacro("NMDC_LiveFilterChooseColumn")
    If Not ActiveSheet Is Nothing Then
        If TypeName(ActiveSheet) = "Worksheet" Then NMDC_LiveFilterRefreshDisplay ActiveSheet
    End If
    On Error GoTo 0
End Sub

Public Sub NMDC_ApplyLiveFilterUX()
    NMDC_LiveFilterWake
End Sub

Private Sub NMDC_LiveFilterInstallSheet(ByVal ws As Worksheet)
    On Error GoTo Handler

    Dim table As ListObject
    Dim area As Range
    Dim displayBox As Shape
    Dim button As Shape

    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Exit Sub

    ' Remove any legacy ActiveX search box from broken/older builds.
    On Error Resume Next
    ws.OLEObjects("TxtBox_Search").Delete
    On Error GoTo Handler

    ' Never merge the filter row. This avoids merge warnings and keeps setup safe.
    On Error Resume Next
    ws.Range("A3:J3").UnMerge
    On Error GoTo Handler
    ws.Range("A3:J3").ClearContents
    ws.Rows(3).RowHeight = 28

    With ws.Range("A3")
        .Value = "LIVE FILTER"
        .Font.Name = "Aptos"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(31, 70, 90)
        .VerticalAlignment = xlCenter
    End With

    On Error Resume Next
    ws.Shapes(NMDC_SEARCH_DISPLAY).Delete
    ws.Shapes(NMDC_SELECT_BUTTON).Delete
    ws.Shapes(NMDC_RESET_BUTTON).Delete
    On Error GoTo Handler

    Set area = ws.Range("B3:E3")
    Set displayBox = ws.Shapes.AddShape(5, area.Left + 1, area.Top + 1, area.Width - 2, area.Height - 2)
    displayBox.Name = NMDC_SEARCH_DISPLAY
    displayBox.OnAction = "NMDC_LiveFilterStartCapture"
    displayBox.Fill.ForeColor.RGB = RGB(255, 255, 255)
    displayBox.Line.ForeColor.RGB = RGB(170, 180, 190)
    displayBox.TextFrame.HorizontalAlignment = xlLeft
    displayBox.TextFrame.VerticalAlignment = 3
    displayBox.TextFrame.Characters.Font.Name = "Aptos"
    displayBox.TextFrame.Characters.Font.Size = 9
    displayBox.TextFrame.Characters.Font.Color = RGB(40, 48, 56)
    displayBox.Placement = xlMoveAndSize

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

    NMDC_LiveFilterRefreshDisplay ws
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_SHEET_SETUP_ERROR", _
        "Excel could not prepare Live Filter on " & ws.Name & ".", _
        Err.Number & " - " & Err.Description
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

Public Sub NMDC_LiveFilterChooseColumn()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim picked As Range
    Dim hit As Range
    Dim targetColumn As ListColumn

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

    NMDC_LiveFilterSetTarget ws, targetColumn
    mQuery = ""
    mCaptureSheetName = ws.Name
    NMDC_LiveFilterApplyQuery ws
    NMDC_LiveFilterStartCapture
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_TARGET_ERROR", _
        "Excel could not change the Live Filter target column.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterStartCapture()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim targetRange As Range

    If ActiveSheet Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    If NMDC_LiveFilterTableForSheet(ws) Is Nothing Then Exit Sub

    Set targetRange = NMDC_LiveFilterTargetRange(ws)
    If targetRange Is Nothing Then
        NMDC_LiveFilterChooseColumn
        Exit Sub
    End If

    If mCaptureActive Then NMDC_LiveFilterStopCapture False
    mCaptureActive = True
    mCaptureSheetName = ws.Name
    NMDC_LiveFilterBindCaptureKeys
    NMDC_LiveFilterRefreshDisplay ws
    Application.StatusBar = "NMDC Live Filter ACTIVE - type to filter " & NMDC_LiveFilterTargetName(ws) & _
                            "; Backspace edits; Delete clears; Esc or Enter stops typing mode."
    Exit Sub

Handler:
    NMDC_LiveFilterStopCapture False
    NMDC_LogError "LIVE_FILTER_CAPTURE_ERROR", _
        "Excel could not start Live Filter typing mode.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterStopCapture(Optional ByVal refreshDisplay As Boolean = True)
    On Error Resume Next
    NMDC_LiveFilterReleaseCaptureKeys
    mCaptureActive = False
    Application.StatusBar = False
    If refreshDisplay Then
        If Not ActiveSheet Is Nothing Then
            If TypeName(ActiveSheet) = "Worksheet" Then NMDC_LiveFilterRefreshDisplay ActiveSheet
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub NMDC_LiveFilterClear()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim table As ListObject
    Dim fieldNumber As Long

    If ActiveSheet Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    Set targetRange = NMDC_LiveFilterTargetRange(ws)

    NMDC_LiveFilterStopCapture False
    mQuery = ""

    If Not targetRange Is Nothing Then
        Set table = targetRange.ListObject
        If Not table Is Nothing Then
            fieldNumber = targetRange.Column - table.Range.Column + 1
            On Error Resume Next
            table.Range.AutoFilter Field:=fieldNumber
            On Error GoTo Handler
        End If
    End If

    NMDC_LiveFilterRefreshDisplay ws
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_CLEAR_ERROR", _
        "Excel could not reset the Live Filter.", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_LiveFilterAppend(ByVal oneChar As String)
    If Not mCaptureActive Then Exit Sub
    If Not NMDC_LiveFilterCaptureContextValid() Then Exit Sub
    mQuery = mQuery & oneChar
    NMDC_LiveFilterApplyQuery ThisWorkbook.Worksheets(mCaptureSheetName)
End Sub

Public Sub NMDC_LF_Backspace()
    If Not mCaptureActive Then Exit Sub
    If Len(mQuery) > 0 Then mQuery = Left$(mQuery, Len(mQuery) - 1)
    If NMDC_LiveFilterCaptureContextValid() Then NMDC_LiveFilterApplyQuery ThisWorkbook.Worksheets(mCaptureSheetName)
End Sub

Public Sub NMDC_LF_ClearQuery()
    If Not mCaptureActive Then Exit Sub
    mQuery = ""
    If NMDC_LiveFilterCaptureContextValid() Then NMDC_LiveFilterApplyQuery ThisWorkbook.Worksheets(mCaptureSheetName)
End Sub

Public Sub NMDC_LF_Stop()
    NMDC_LiveFilterStopCapture True
End Sub

Private Function NMDC_LiveFilterCaptureContextValid() As Boolean
    On Error GoTo Failed
    If Not mCaptureActive Then Exit Function
    If ActiveWorkbook Is Nothing Then GoTo Failed
    If Not ActiveWorkbook Is ThisWorkbook Then GoTo Failed
    If ActiveSheet Is Nothing Then GoTo Failed
    If StrComp(CStr(ActiveSheet.Name), mCaptureSheetName, vbTextCompare) <> 0 Then GoTo Failed
    NMDC_LiveFilterCaptureContextValid = True
    Exit Function
Failed:
    NMDC_LiveFilterStopCapture False
End Function

Private Sub NMDC_LiveFilterApplyQuery(ByVal ws As Worksheet)
    On Error GoTo Handler

    Dim targetRange As Range
    Dim table As ListObject
    Dim fieldNumber As Long
    Dim cleanText As String
    Dim exactMode As Boolean

    Set targetRange = NMDC_LiveFilterTargetRange(ws)
    If targetRange Is Nothing Then Exit Sub
    Set table = targetRange.ListObject
    If table Is Nothing Then Exit Sub

    cleanText = mQuery
    exactMode = False
    If Len(cleanText) > 1 Then
        If Left$(cleanText, 1) = Chr$(34) And Right$(cleanText, 1) = Chr$(34) Then
            exactMode = True
            cleanText = Mid$(cleanText, 2, Len(cleanText) - 2)
        End If
    End If

    fieldNumber = targetRange.Column - table.Range.Column + 1
    If fieldNumber < 1 Or fieldNumber > table.ListColumns.Count Then Exit Sub

    Application.ScreenUpdating = False
    On Error Resume Next
    table.Range.AutoFilter Field:=fieldNumber
    On Error GoTo Handler

    If Len(cleanText) > 0 Then
        If exactMode Then
            ' Exact whole-cell match. Native AutoFilter is intentionally used for speed.
            table.Range.AutoFilter Field:=fieldNumber, Criteria1:=cleanText
        Else
            table.Range.AutoFilter Field:=fieldNumber, _
                Criteria1:="=*" & NMDC_LiveFilterEscapeWildcards(cleanText) & "*"
        End If
    End If

    Application.ScreenUpdating = True
    NMDC_LiveFilterRefreshDisplay ws
    Exit Sub

Handler:
    Application.ScreenUpdating = True
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

    If targetColumn.DataBodyRange Is Nothing Then Exit Sub
    ThisWorkbook.Names.Add Name:=anchorName, RefersTo:=targetColumn.DataBodyRange, Visible:=False
End Sub

Private Function NMDC_LiveFilterTargetRange(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set NMDC_LiveFilterTargetRange = ThisWorkbook.Names(NMDC_LiveFilterAnchorName(ws)).RefersToRange
    On Error GoTo 0
End Function

Private Function NMDC_LiveFilterAnchorName(ByVal ws As Worksheet) As String
    NMDC_LiveFilterAnchorName = NMDC_ANCHOR_PREFIX & ws.CodeName
End Function

Private Function NMDC_LiveFilterTargetName(ByVal ws As Worksheet) As String
    On Error GoTo Missing
    Dim targetRange As Range
    Set targetRange = NMDC_LiveFilterTargetRange(ws)
    If targetRange Is Nothing Then GoTo Missing
    NMDC_LiveFilterTargetName = CStr(targetRange.ListObject.HeaderRowRange.Cells( _
        1, targetRange.Column - targetRange.ListObject.Range.Column + 1).Value)
    Exit Function
Missing:
    NMDC_LiveFilterTargetName = "selected column"
End Function

Private Sub NMDC_LiveFilterRefreshDisplay(ByVal ws As Worksheet)
    On Error Resume Next
    Dim displayBox As Shape
    Dim selectButton As Shape
    Dim targetName As String
    Dim textValue As String

    Set displayBox = ws.Shapes(NMDC_SEARCH_DISPLAY)
    Set selectButton = ws.Shapes(NMDC_SELECT_BUTTON)
    targetName = NMDC_LiveFilterTargetName(ws)

    If NMDC_LiveFilterTargetRange(ws) Is Nothing Then
        textValue = "Select a column, then type."
        If Not selectButton Is Nothing Then selectButton.TextFrame.Characters.Text = "SELECT COLUMN"
    ElseIf mCaptureActive And StrComp(ws.Name, mCaptureSheetName, vbTextCompare) = 0 Then
        textValue = "SEARCH " & targetName & ": " & mQuery & "  |  LIVE - Esc stops"
        If Not selectButton Is Nothing Then selectButton.TextFrame.Characters.Text = "COLUMN: " & targetName
    Else
        textValue = "SEARCH " & targetName & ": " & mQuery & "  |  click here to type"
        If Not selectButton Is Nothing Then selectButton.TextFrame.Characters.Text = "COLUMN: " & targetName
    End If

    If Not displayBox Is Nothing Then displayBox.TextFrame.Characters.Text = textValue
    On Error GoTo 0
End Sub

Public Sub NMDC_LiveFilterSheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Retained for existing workbook event wiring. Live typing is handled by OnKey.
End Sub

Public Sub NMDC_LiveFilterSheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If mCaptureActive Then
        If StrComp(CStr(Sh.Name), mCaptureSheetName, vbTextCompare) <> 0 Then NMDC_LiveFilterStopCapture False
    End If
    If NMDC_LiveFilterTableForSheet(Sh) Is Nothing Then Exit Sub
    NMDC_LiveFilterRefreshDisplay Sh
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

Private Function NMDC_LiveFilterQualifiedMacro(ByVal macroName As String) As String
    NMDC_LiveFilterQualifiedMacro = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

Private Sub NMDC_LiveFilterBind(ByVal keyText As String, ByVal macroName As String)
    Application.OnKey keyText, NMDC_LiveFilterQualifiedMacro(macroName)
End Sub

Private Sub NMDC_LiveFilterBindCaptureKeys()
    Dim letters As Variant
    Dim digits As Variant
    Dim i As Long
    Dim ch As String

    letters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", _
                    "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    digits = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")

    For i = LBound(letters) To UBound(letters)
        ch = LCase$(CStr(letters(i)))
        NMDC_LiveFilterBind ch, "NMDC_LF_" & CStr(letters(i))
        NMDC_LiveFilterBind "+" & ch, "NMDC_LF_" & CStr(letters(i))
    Next i
    For i = LBound(digits) To UBound(digits)
        NMDC_LiveFilterBind CStr(digits(i)), "NMDC_LF_" & CStr(digits(i))
    Next i

    NMDC_LiveFilterBind " ", "NMDC_LF_Space"
    NMDC_LiveFilterBind "-", "NMDC_LF_Hyphen"
    NMDC_LiveFilterBind "+-", "NMDC_LF_Underscore"
    NMDC_LiveFilterBind ".", "NMDC_LF_Dot"
    NMDC_LiveFilterBind "/", "NMDC_LF_Slash"
    NMDC_LiveFilterBind "{+}", "NMDC_LF_Plus"
    NMDC_LiveFilterBind "+'", "NMDC_LF_Quote"
    NMDC_LiveFilterBind "{BACKSPACE}", "NMDC_LF_Backspace"
    NMDC_LiveFilterBind "{DELETE}", "NMDC_LF_ClearQuery"
    NMDC_LiveFilterBind "{ESC}", "NMDC_LF_Stop"
    NMDC_LiveFilterBind "~", "NMDC_LF_Stop"
    NMDC_LiveFilterBind "{ENTER}", "NMDC_LF_Stop"
    NMDC_LiveFilterBind "{TAB}", "NMDC_LF_Stop"
End Sub

Private Sub NMDC_LiveFilterReleaseCaptureKeys()
    Dim letters As Variant
    Dim digits As Variant
    Dim i As Long
    Dim ch As String

    letters = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", _
                    "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
    digits = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")

    On Error Resume Next
    For i = LBound(letters) To UBound(letters)
        ch = CStr(letters(i))
        Application.OnKey ch
        Application.OnKey "+" & ch
    Next i
    For i = LBound(digits) To UBound(digits)
        Application.OnKey CStr(digits(i))
    Next i
    Application.OnKey " "
    Application.OnKey "-"
    Application.OnKey "+-"
    Application.OnKey "."
    Application.OnKey "/"
    Application.OnKey "{+}"
    Application.OnKey "+'"
    Application.OnKey "{BACKSPACE}"
    Application.OnKey "{DELETE}"
    Application.OnKey "{ESC}"
    Application.OnKey "~"
    Application.OnKey "{ENTER}"
    Application.OnKey "{TAB}"
    On Error GoTo 0
End Sub

' Character handlers - kept public because Application.OnKey can only call public no-argument procedures.
Public Sub NMDC_LF_A(): NMDC_LiveFilterAppend "a": End Sub
Public Sub NMDC_LF_B(): NMDC_LiveFilterAppend "b": End Sub
Public Sub NMDC_LF_C(): NMDC_LiveFilterAppend "c": End Sub
Public Sub NMDC_LF_D(): NMDC_LiveFilterAppend "d": End Sub
Public Sub NMDC_LF_E(): NMDC_LiveFilterAppend "e": End Sub
Public Sub NMDC_LF_F(): NMDC_LiveFilterAppend "f": End Sub
Public Sub NMDC_LF_G(): NMDC_LiveFilterAppend "g": End Sub
Public Sub NMDC_LF_H(): NMDC_LiveFilterAppend "h": End Sub
Public Sub NMDC_LF_I(): NMDC_LiveFilterAppend "i": End Sub
Public Sub NMDC_LF_J(): NMDC_LiveFilterAppend "j": End Sub
Public Sub NMDC_LF_K(): NMDC_LiveFilterAppend "k": End Sub
Public Sub NMDC_LF_L(): NMDC_LiveFilterAppend "l": End Sub
Public Sub NMDC_LF_M(): NMDC_LiveFilterAppend "m": End Sub
Public Sub NMDC_LF_N(): NMDC_LiveFilterAppend "n": End Sub
Public Sub NMDC_LF_O(): NMDC_LiveFilterAppend "o": End Sub
Public Sub NMDC_LF_P(): NMDC_LiveFilterAppend "p": End Sub
Public Sub NMDC_LF_Q(): NMDC_LiveFilterAppend "q": End Sub
Public Sub NMDC_LF_R(): NMDC_LiveFilterAppend "r": End Sub
Public Sub NMDC_LF_S(): NMDC_LiveFilterAppend "s": End Sub
Public Sub NMDC_LF_T(): NMDC_LiveFilterAppend "t": End Sub
Public Sub NMDC_LF_U(): NMDC_LiveFilterAppend "u": End Sub
Public Sub NMDC_LF_V(): NMDC_LiveFilterAppend "v": End Sub
Public Sub NMDC_LF_W(): NMDC_LiveFilterAppend "w": End Sub
Public Sub NMDC_LF_X(): NMDC_LiveFilterAppend "x": End Sub
Public Sub NMDC_LF_Y(): NMDC_LiveFilterAppend "y": End Sub
Public Sub NMDC_LF_Z(): NMDC_LiveFilterAppend "z": End Sub
Public Sub NMDC_LF_0(): NMDC_LiveFilterAppend "0": End Sub
Public Sub NMDC_LF_1(): NMDC_LiveFilterAppend "1": End Sub
Public Sub NMDC_LF_2(): NMDC_LiveFilterAppend "2": End Sub
Public Sub NMDC_LF_3(): NMDC_LiveFilterAppend "3": End Sub
Public Sub NMDC_LF_4(): NMDC_LiveFilterAppend "4": End Sub
Public Sub NMDC_LF_5(): NMDC_LiveFilterAppend "5": End Sub
Public Sub NMDC_LF_6(): NMDC_LiveFilterAppend "6": End Sub
Public Sub NMDC_LF_7(): NMDC_LiveFilterAppend "7": End Sub
Public Sub NMDC_LF_8(): NMDC_LiveFilterAppend "8": End Sub
Public Sub NMDC_LF_9(): NMDC_LiveFilterAppend "9": End Sub
Public Sub NMDC_LF_Space(): NMDC_LiveFilterAppend " ": End Sub
Public Sub NMDC_LF_Hyphen(): NMDC_LiveFilterAppend "-": End Sub
Public Sub NMDC_LF_Underscore(): NMDC_LiveFilterAppend "_": End Sub
Public Sub NMDC_LF_Dot(): NMDC_LiveFilterAppend ".": End Sub
Public Sub NMDC_LF_Slash(): NMDC_LiveFilterAppend "/": End Sub
Public Sub NMDC_LF_Plus(): NMDC_LiveFilterAppend "+": End Sub
Public Sub NMDC_LF_Quote(): NMDC_LiveFilterAppend Chr$(34): End Sub
