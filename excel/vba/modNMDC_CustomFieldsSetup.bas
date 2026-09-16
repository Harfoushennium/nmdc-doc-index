Attribute VB_Name = "modNMDC_CustomFieldsSetup"
Option Explicit

Public Sub NMDC_EnsureCustomFieldsStructure()
    On Error GoTo Handler

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Custom Fields")
    On Error GoTo Handler

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Custom Fields"
    End If

    ws.Range("A1").Value = "Custom Fields & Keywords"
    ws.Range("A2").Value = "Create user-defined derived columns in Master Documents without changing the canonical NMDC index data."
    ws.Range("A1").Font.Name = "Aptos"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Font.Name = "Aptos"
    ws.Range("A2").Font.Size = 10

    NMDC_EnsureCustomTable ws, "CustomFields", 5, 1, _
        Array("Enabled?", "Field Name", "Search In", "Match Behavior", "Separator", "Notes")
    NMDC_EnsureCustomTable ws, "KeywordMappings", 5, 8, _
        Array("Enabled?", "Field Name", "Keyword / Pattern", "Result", "Match Type", "Priority", "Notes")

    NMDC_EnsureCustomFieldsHomeButton
    NMDC_InstallCustomFieldSelectionEvent
    Exit Sub

Handler:
    NMDC_LogError "CUSTOM_FIELDS_STRUCTURE_ERROR", _
        "Excel could not create or validate the Custom Fields & Keywords workspace.", _
        Err.Number & " - " & Err.Description
End Sub

Private Sub NMDC_EnsureCustomTable( _
    ByVal ws As Worksheet, ByVal tableName As String, ByVal headerRow As Long, _
    ByVal firstColumn As Long, ByVal headers As Variant)

    Dim table As ListObject
    Dim sourceRange As Range
    Dim headerCount As Long
    Dim i As Long

    Set table = Nothing
    On Error Resume Next
    Set table = ws.ListObjects(tableName)
    On Error GoTo 0

    headerCount = UBound(headers) - LBound(headers) + 1
    ws.Range(ws.Cells(headerRow, firstColumn), _
             ws.Cells(headerRow + 1, firstColumn + headerCount - 1)).UnMerge

    For i = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, firstColumn + i - LBound(headers)).Value = headers(i)
    Next i

    If table Is Nothing Then
        Set sourceRange = ws.Range( _
            ws.Cells(headerRow, firstColumn), _
            ws.Cells(headerRow + 1, firstColumn + headerCount - 1))
        Set table = ws.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
        table.Name = tableName
        table.TableStyle = "TableStyleMedium2"
    Else
        Set sourceRange = ws.Range( _
            ws.Cells(headerRow, firstColumn), _
            ws.Cells(headerRow + Application.Max(1, table.ListRows.Count), firstColumn + headerCount - 1))
        table.Resize sourceRange
    End If

    If table.ListRows.Count = 0 Then table.ListRows.Add
End Sub

Public Sub NMDC_EnsureCustomFieldsHomeButton()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim area As Range
    Dim button As Shape

    Set ws = ThisWorkbook.Worksheets("Home")
    On Error Resume Next
    ws.Shapes("NMDC_Action_20").Delete
    On Error GoTo Handler

    Set area = ws.Range("E40:H41")
    area.Hyperlinks.Delete
    area.Cells(1, 1).Value = "Custom Fields & Keywords"

    Set button = ws.Shapes.AddShape(5, area.Left + 2, area.Top + 2, area.Width - 4, area.Height - 4)
    button.Name = "NMDC_Action_20"
    button.OnAction = "NMDC_OpenCustomFields"
    button.TextFrame.Characters.Text = "Custom Fields & Keywords" & vbLf & "Add derived columns from keyword lists"
    button.TextFrame.HorizontalAlignment = xlCenter
    button.TextFrame.VerticalAlignment = xlCenter
    button.Fill.ForeColor.RGB = RGB(20, 108, 148)
    button.Line.ForeColor.RGB = RGB(20, 108, 148)
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 8.5
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    button.AlternativeText = "Custom Fields & Keywords: add user-defined Master Documents columns and populate them from editable keyword mappings."
    button.Placement = xlMoveAndSize
    Exit Sub

Handler:
    NMDC_LogError "CUSTOM_FIELDS_HOME_BUTTON_ERROR", _
        "Excel could not add the Custom Fields & Keywords Home button.", _
        Err.Number & " - " & Err.Description
End Sub

' Core CSV refresh intentionally owns only canonical Master Documents columns.
' If it trims the table back to those core columns, this lightweight check
' notices the missing enabled custom field and rebuilds the enrichment layer
' the next time Master Documents is activated or clicked.
Public Sub NMDC_EnsureCustomFieldsCurrent(ByVal Sh As Object)
    On Error GoTo SoftFail

    Dim ws As Worksheet
    Dim definitions As ListObject
    Dim master As ListObject
    Dim row As ListRow
    Dim fieldName As String
    Dim enabledValue As Variant
    Dim target As ListColumn

    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Set ws = Sh
    If StrComp(ws.Name, "Master Documents", vbTextCompare) <> 0 Then Exit Sub

    Set definitions = ThisWorkbook.Worksheets("Custom Fields").ListObjects("CustomFields")
    Set master = ws.ListObjects("MasterDocuments")
    If definitions.DataBodyRange Is Nothing Then Exit Sub

    For Each row In definitions.ListRows
        fieldName = Trim$(CStr(row.Range.Cells(1, definitions.ListColumns("Field Name").Index).Value))
        enabledValue = row.Range.Cells(1, definitions.ListColumns("Enabled?").Index).Value
        If Len(fieldName) > 0 And NMDC_CustomSetupChecked(enabledValue) Then
            Set target = Nothing
            On Error Resume Next
            Set target = master.ListColumns(fieldName)
            On Error GoTo SoftFail
            If target Is Nothing Then
                Application.Run "NMDC_ApplyCustomFieldsToMaster", False
                Exit Sub
            End If
        End If
    Next row
    Exit Sub

SoftFail:
    ' This guard must never interrupt normal worksheet navigation.
End Sub

' Install a separate workbook selection-change event. This complements the live
' filter SheetChange/SheetActivate handlers without needing to edit those event
' procedures. It makes a post-refresh custom field restore effectively automatic:
' the first click in Master Documents performs only a lightweight definition check,
' and rebuilds the derived columns only when one is actually missing.
Private Sub NMDC_InstallCustomFieldSelectionEvent()
    On Error GoTo SoftFail

    Dim project As Object
    Dim component As Object
    Dim codeModule As Object
    Dim sourceText As String
    Dim eventCode As String

    Set project = Nothing
    On Error Resume Next
    Set project = ThisWorkbook.VBProject
    On Error GoTo SoftFail
    If project Is Nothing Then Exit Sub

    Set component = project.VBComponents(ThisWorkbook.CodeName)
    Set codeModule = component.CodeModule
    sourceText = ""
    If codeModule.CountOfLines > 0 Then
        sourceText = codeModule.Lines(1, codeModule.CountOfLines)
    End If

    If InStr(1, sourceText, "Private Sub Workbook_SheetSelectionChange", vbTextCompare) > 0 Then Exit Sub

    eventCode = vbCrLf & _
        "Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf & _
        "    On Error Resume Next" & vbCrLf & _
        "    NMDC_EnsureCustomFieldsCurrent Sh" & vbCrLf & _
        "    On Error GoTo 0" & vbCrLf & _
        "End Sub" & vbCrLf
    codeModule.AddFromString eventCode
    Exit Sub

SoftFail:
    ' Best-effort only. The one-time production setup normally grants access.
End Sub

Private Function NMDC_CustomSetupChecked(ByVal value As Variant) As Boolean
    If VarType(value) = vbBoolean Then
        NMDC_CustomSetupChecked = CBool(value)
        Exit Function
    End If

    Select Case UCase$(Trim$(CStr(value)))
        Case "TRUE", "YES", "Y", "1", "ON", "CHECKED"
            NMDC_CustomSetupChecked = True
        Case Else
            NMDC_CustomSetupChecked = False
    End Select
End Function
