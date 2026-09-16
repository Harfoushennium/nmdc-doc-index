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
