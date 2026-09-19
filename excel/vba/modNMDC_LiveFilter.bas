Attribute VB_Name = "modNMDC_LiveFilter"
Option Explicit

Private Const NMDC_LIVE_HELPER As String = "__NMDC_LiveFilter"
Private Const NMDC_LIVE_ALL_COLUMNS As String = "ALL COLUMNS"
Private Const NMDC_LIVE_OPEN_BUTTON As String = "NMDC_LiveFilter_Open"
Private Const NMDC_LIVE_RESET_BUTTON As String = "NMDC_LiveFilter_Reset"

' Dynamic filter concept adapted from the owner's forensic Live Filter tool.
' The production workbook deliberately uses worksheet cells instead of an
' ActiveX textbox so the feature remains reliable on locked-down corporate PCs.
' The modeless MSForms input supplies the visible caret and per-keystroke
' Change event.  B3 mirrors the current query for persistence and filtering;
' it is not presented as the live typing surface.
'
' Syntax:
'   word1 word2   -> AND (both must appear)
'   word1+word2   -> AND (same as spaces)
'   -word         -> exclude rows containing the word
'   "Exact Text"  -> exact, case-sensitive cell match
' The Select Column action chooses one table column; ALL COLUMNS searches the
' whole row.  Wildcards * and ? are supported in terms.

Public Sub NMDC_LiveFilterInitialize()
    On Error GoTo Handler
    NMDC_ApplyLiveFilterUX
    Exit Sub
Handler:
    NMDC_LogError "LIVE_FILTER_INIT_ERROR", _
        "Excel could not initialize the live filter controls.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ApplyLiveFilterUX()
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
        If Not ws Is Nothing Then NMDC_LiveFilterConfigureSheet ws
    Next item
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_UX_ERROR", _
        "Excel could not prepare one or more live filter search bars.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterWake()
    ' Compatibility entry point used by Pending Update rebuilds.  It only
    ' refreshes the form launch UI; it never installs a global key hook.
    NMDC_ApplyLiveFilterUX
End Sub

Private Sub NMDC_LiveFilterConfigureSheet(ByVal ws As Worksheet)
    On Error GoTo Handler

    Dim table As ListObject
    Dim searchText As String
    Dim targetText As String
    Dim area As Range
    Dim button As Shape

    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Exit Sub

    searchText = CStr(ws.Range("B3").Value)
    targetText = CStr(ws.Range("F3").Value)

    ws.Range("A3").Value = "LIVE FILTER"
    ' B3/F3 are persisted mirrors for the modeless form state; the form owns
    ' the editable caret and per-keystroke Change event.
    ws.Range("B3:D3").UnMerge
    ws.Range("F3:H3").UnMerge
    ws.Range("B3:D3").ClearContents
    ws.Range("F3:H3").ClearContents
    ws.Range("B3").Value = searchText
    ws.Range("E3").Value = "IN"
    If Len(Trim$(targetText)) = 0 Then targetText = NMDC_LIVE_ALL_COLUMNS
    ws.Range("F3").Value = targetText

    With ws.Range("A3:H3")
        .Font.Name = "Aptos"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .RowHeight = 24
    End With
    With ws.Range("A3")
        .Font.Bold = True
        .Font.Color = RGB(31, 70, 90)
    End With
    With ws.Range("E3")
        .Font.Bold = True
        .Font.Color = RGB(31, 70, 90)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("B3:D3")
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(20, 108, 148)
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft
    End With
    With ws.Range("F3:H3")
        .Interior.Color = RGB(232, 241, 247)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(20, 108, 148)
        .Font.Color = RGB(31, 70, 90)
        .Font.Bold = True
    End With

    On Error Resume Next
    ws.Range("B3").Validation.Delete
    ws.Range("B3").Validation.Add Type:=xlValidateInputOnly
    ws.Range("B3").Validation.InputTitle = "Live Filter"
    ws.Range("B3").Validation.InputMessage = _
        "Click OPEN LIVE FILTER to type; spaces or + = AND; -word = exclude; quotes = exact; * and ? = wildcards. Use Choose Column to change the target."
    ws.Range("B3").Validation.ShowInput = True
    On Error GoTo Handler

    On Error Resume Next
    ws.Shapes(NMDC_LIVE_OPEN_BUTTON).Delete
    ws.Shapes(NMDC_LIVE_RESET_BUTTON).Delete
    On Error GoTo Handler
    Set area = ws.Range("I3:J3")
    Set button = ws.Shapes.AddShape(5, area.Left + 1, area.Top + 1, area.Width - 2, area.Height - 2)
    button.Name = NMDC_LIVE_OPEN_BUTTON
    button.OnAction = "NMDC_LiveFilterShow"
    button.TextFrame.Characters.Text = "OPEN LIVE FILTER"
    NMDC_LiveFilterFormatButton button, RGB(20, 108, 148)
    Set area = ws.Range("K3:L3")
    Set button = ws.Shapes.AddShape(5, area.Left + 1, area.Top + 1, area.Width - 2, area.Height - 2)
    button.Name = NMDC_LIVE_RESET_BUTTON
    button.OnAction = "NMDC_LiveFilterClear"
    button.TextFrame.Characters.Text = "RESET"
    NMDC_LiveFilterFormatButton button, RGB(91, 100, 112)

    NMDC_LiveFilterEnsureHelper table

    If Len(Trim$(searchText)) > 0 Then NMDC_LiveFilterApplyForSheet ws
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_SHEET_UX_ERROR", _
        "Excel could not prepare the live filter on " & ws.Name & ".", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_LiveFilterFormatButton(ByVal button As Shape, ByVal fillColor As Long)
    With button
        .TextFrame.HorizontalAlignment = xlCenter
        .TextFrame.VerticalAlignment = 3
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = fillColor
        .TextFrame.Characters.Font.Name = "Aptos"
        .TextFrame.Characters.Font.Size = 8
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Placement = xlMoveAndSize
    End With
End Sub

Public Sub NMDC_LiveFilterSheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject

    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Set ws = Sh
    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Exit Sub

    If Intersect(Target, ws.Range("B3:D3")) Is Nothing And _
       Intersect(Target, ws.Range("F3:H3")) Is Nothing Then Exit Sub

    NMDC_LiveFilterApplyForSheet ws
    Exit Sub

Handler:
    NMDC_LogError "LIVE_FILTER_CHANGE_ERROR", _
        "Excel could not apply the live filter on " & Sh.Name & ".", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterSheetActivate(ByVal Sh As Object)
    On Error GoTo Handler
    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If NMDC_LiveFilterTableForSheet(Sh) Is Nothing Then Exit Sub
    NMDC_LiveFilterConfigureSheet Sh
    Exit Sub
Handler:
    NMDC_LogError "LIVE_FILTER_ACTIVATE_ERROR", _
        "Excel could not refresh the live filter controls.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterChooseColumn()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim picked As Range
    Dim hit As Range
    Dim headerName As String
    Dim previousEvents As Boolean

    If ActiveSheet Is Nothing Then Exit Sub
    Set ws = ActiveSheet
    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then
        MsgBox "Live Filter is available on the main data/review tables only.", _
               vbInformation, "NMDC Document Index"
        Exit Sub
    End If

    On Error Resume Next
    Set picked = Application.InputBox( _
        "Click the HEADER of the column to search." & vbCrLf & vbCrLf & _
        "Cancel to keep the current target. To search the whole row, type ALL COLUMNS in the target box.", _
        "Choose Live Filter Column", Type:=8)
    On Error GoTo Handler
    If picked Is Nothing Then Exit Sub

    Set hit = Intersect(picked.Cells(1, 1), table.HeaderRowRange)
    If hit Is Nothing Then
        MsgBox "Please click a header cell inside the table on this sheet.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    headerName = CStr(hit.Value)
    If Left$(headerName, 7) = "__NMDC_" Then
        MsgBox "That is an internal helper column. Choose a normal user-visible field.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range("F3").Value = headerName
    Application.EnableEvents = previousEvents

    NMDC_LiveFilterApplyForSheet ws
    On Error Resume Next
    frmNMDC_LiveFilter.NMDC_Bind ws.Name, CStr(ws.Range("B3").Value), headerName
    On Error GoTo 0
    Exit Sub

Handler:
    On Error Resume Next
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    NMDC_LogError "LIVE_FILTER_TARGET_ERROR", _
        "Excel could not change the live filter target column.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterShow()
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim table As ListObject
    If ActiveSheet Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Exit Sub
    NMDC_LiveFilterConfigureSheet ws
    frmNMDC_LiveFilter.NMDC_Bind ws.Name, CStr(ws.Range("B3").Value), CStr(ws.Range("F3").Value)
    frmNMDC_LiveFilter.Show vbModeless
    Exit Sub
Handler:
    NMDC_LogError "LIVE_FILTER_FORM_ERROR", _
        "Excel could not open the modeless Live Filter input.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterFormChanged(ByVal sheetName As String, ByVal queryText As String)
    On Error GoTo Handler
    Dim ws As Worksheet
    Dim previousEvents As Boolean
    Set ws = ThisWorkbook.Worksheets(sheetName)
    previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range("B3").Value = queryText
    Application.EnableEvents = previousEvents
    NMDC_LiveFilterApplyForSheet ws
    Exit Sub
Handler:
    On Error Resume Next
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    NMDC_LogError "LIVE_FILTER_CHANGE_ERROR", _
        "Excel could not apply the Live Filter while typing.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterClear()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim previousEvents As Boolean

    If ActiveSheet Is Nothing Then Exit Sub
    Set ws = ActiveSheet
    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Exit Sub

    previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range("B3").Value = ""
    Application.EnableEvents = previousEvents
    NMDC_LiveFilterApplyForSheet ws
    Exit Sub

Handler:
    On Error Resume Next
    Application.EnableEvents = previousEvents
    On Error GoTo 0
    NMDC_LogError "LIVE_FILTER_CLEAR_ERROR", _
        "Excel could not clear the live filter.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LiveFilterApplyForSheet(ByVal ws As Worksheet)
    On Error GoTo Handler

    Dim table As ListObject
    Dim helper As ListColumn
    Dim searchText As String
    Dim targetName As String
    Dim targetIndex As Long
    Dim values As Variant
    Dim resultValues() As Variant
    Dim rowCount As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim rowText As String
    Dim cleanText As String
    Dim strictMode As Boolean
    Dim rawTokens() As String
    Dim posTerms() As String
    Dim negTerms() As String
    Dim pCount As Long
    Dim nCount As Long
    Dim token As String
    Dim i As Long
    Dim matched As Boolean
    Dim previousScreen As Boolean
    Dim previousEvents As Boolean

    Set table = NMDC_LiveFilterTableForSheet(ws)
    If table Is Nothing Then Exit Sub
    Set helper = NMDC_LiveFilterEnsureHelper(table)
    If helper Is Nothing Then Exit Sub

    searchText = Trim$(CStr(ws.Range("B3").Value))
    targetName = Trim$(CStr(ws.Range("F3").Value))
    If Len(targetName) = 0 Then targetName = NMDC_LIVE_ALL_COLUMNS

    If Len(searchText) = 0 Then
        On Error Resume Next
        table.Range.AutoFilter Field:=helper.Index
        helper.Range.EntireColumn.Hidden = True
        On Error GoTo Handler
        Exit Sub
    End If

    If table.DataBodyRange Is Nothing Then Exit Sub
    rowCount = table.DataBodyRange.Rows.Count
    values = table.DataBodyRange.Value2
    ReDim resultValues(1 To rowCount, 1 To 1)

    targetIndex = 0
    If StrComp(targetName, NMDC_LIVE_ALL_COLUMNS, vbTextCompare) <> 0 Then
        On Error Resume Next
        targetIndex = table.ListColumns(targetName).Index
        On Error GoTo Handler
        If targetIndex > 0 Then
            If Left$(CStr(table.ListColumns(targetIndex).Name), 7) = "__NMDC_" Then targetIndex = 0
        End If
    End If

    cleanText = searchText
    strictMode = False
    If Len(cleanText) > 1 Then
        If Left$(cleanText, 1) = Chr$(34) And Right$(cleanText, 1) = Chr$(34) Then
            strictMode = True
            cleanText = Mid$(cleanText, 2, Len(cleanText) - 2)
        End If
    End If

    pCount = 0
    nCount = 0
    If Not strictMode Then
        rawTokens = Split(Replace(cleanText, "+", " "), " ")
        ReDim posTerms(0 To UBound(rawTokens))
        ReDim negTerms(0 To UBound(rawTokens))
        For i = LBound(rawTokens) To UBound(rawTokens)
            token = Trim$(rawTokens(i))
            If Len(token) > 0 Then
                If Left$(token, 1) = "-" And Len(token) > 1 Then
                    negTerms(nCount) = Mid$(token, 2)
                    nCount = nCount + 1
                Else
                    posTerms(pCount) = token
                    pCount = pCount + 1
                End If
            End If
        Next i
    End If

    previousScreen = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For rowIndex = 1 To rowCount
        matched = True

        If strictMode Then
            matched = False
            If targetIndex > 0 Then
                matched = NMDC_LiveFilterMatchesExact(CStr(values(rowIndex, targetIndex)), cleanText)
            Else
                For columnIndex = 1 To table.ListColumns.Count
                    If Left$(CStr(table.ListColumns(columnIndex).Name), 7) <> "__NMDC_" Then
                        If NMDC_LiveFilterMatchesExact(CStr(values(rowIndex, columnIndex)), cleanText) Then
                            matched = True
                            Exit For
                        End If
                    End If
                Next columnIndex
            End If
        Else
            If targetIndex > 0 Then
                rowText = CStr(values(rowIndex, targetIndex))
            Else
                rowText = NMDC_LiveFilterRowText(table, values, rowIndex)
            End If

            If pCount > 0 Then
                For i = 0 To pCount - 1
                    If Not NMDC_LiveFilterMatchesTerm(rowText, posTerms(i)) Then
                        matched = False
                        Exit For
                    End If
                Next i
            End If

            If matched And nCount > 0 Then
                For i = 0 To nCount - 1
                    If NMDC_LiveFilterMatchesTerm(rowText, negTerms(i)) Then
                        matched = False
                        Exit For
                    End If
                Next i
            End If
        End If

        resultValues(rowIndex, 1) = matched
    Next rowIndex

    helper.DataBodyRange.Value2 = resultValues
    table.Range.AutoFilter Field:=helper.Index, Criteria1:=True
    helper.Range.EntireColumn.Hidden = True

CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    Exit Sub

Handler:
    On Error Resume Next
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    On Error GoTo 0
    NMDC_LogError "LIVE_FILTER_APPLY_ERROR", _
        "Excel could not apply the live filter on " & ws.Name & ".", _
        Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_LiveFilterMatchesExact(ByVal valueText As String, ByVal patternText As String) As Boolean
    If InStr(1, patternText, "*", vbBinaryCompare) > 0 Or _
       InStr(1, patternText, "?", vbBinaryCompare) > 0 Then
        NMDC_LiveFilterMatchesExact = (LCase$(valueText) Like LCase$(patternText))
    Else
        NMDC_LiveFilterMatchesExact = (StrComp(valueText, patternText, vbBinaryCompare) = 0)
    End If
End Function

Private Function NMDC_LiveFilterMatchesTerm(ByVal valueText As String, ByVal termText As String) As Boolean
    If InStr(1, termText, "*", vbBinaryCompare) > 0 Or _
       InStr(1, termText, "?", vbBinaryCompare) > 0 Then
        NMDC_LiveFilterMatchesTerm = (LCase$(valueText) Like LCase$(termText))
    Else
        NMDC_LiveFilterMatchesTerm = (InStr(1, valueText, termText, vbTextCompare) > 0)
    End If
End Function

Private Function NMDC_LiveFilterRowText(ByVal table As ListObject, ByVal values As Variant, ByVal rowIndex As Long) As String
    Dim columnIndex As Long
    Dim valueText As String
    Dim combined As String

    combined = ""
    For columnIndex = 1 To table.ListColumns.Count
        If Left$(CStr(table.ListColumns(columnIndex).Name), 7) <> "__NMDC_" Then
            valueText = CStr(values(rowIndex, columnIndex))
            If Len(valueText) > 0 Then
                If Len(combined) > 0 Then combined = combined & " | "
                combined = combined & valueText
            End If
        End If
    Next columnIndex
    NMDC_LiveFilterRowText = combined
End Function

Public Function NMDC_LiveFilterEnsureHelper(ByVal table As ListObject) As ListColumn
    On Error GoTo Handler

    Dim helper As ListColumn
    Set helper = Nothing
    On Error Resume Next
    Set helper = table.ListColumns(NMDC_LIVE_HELPER)
    On Error GoTo Handler

    If helper Is Nothing Then
        Set helper = table.ListColumns.Add
        helper.Name = NMDC_LIVE_HELPER
    End If

    helper.Range.EntireColumn.Hidden = True
    Set NMDC_LiveFilterEnsureHelper = helper
    Exit Function

Handler:
    NMDC_LogError "LIVE_FILTER_HELPER_ERROR", _
        "Excel could not prepare the internal live-filter helper field for table " & table.Name & ".", _
        Err.Number & " - " & Err.Description
    Set NMDC_LiveFilterEnsureHelper = Nothing
End Function

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
