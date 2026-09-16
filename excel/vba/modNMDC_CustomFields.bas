Attribute VB_Name = "modNMDC_CustomFields"
Option Explicit

Private Const NMDC_CUSTOM_SHEET As String = "Custom Fields"
Private Const NMDC_FIELDS_TABLE As String = "CustomFields"
Private Const NMDC_MAPPINGS_TABLE As String = "KeywordMappings"
Private Const NMDC_FIELD_CHECK_PREFIX As String = "NMDC_CustomFieldCheck_"
Private Const NMDC_MAP_CHECK_PREFIX As String = "NMDC_KeywordMapCheck_"
Private Const NMDC_CUSTOM_FIELDS_FILE As String = "user_custom_fields.csv"
Private Const NMDC_KEYWORD_MAPPINGS_FILE As String = "user_keyword_mappings.csv"

' User-defined enrichment layer for Master Documents.
'
' The canonical engine dataset remains unchanged. Users define derived fields in
' CustomFields and keyword dictionaries in KeywordMappings. The VBA layer then
' adds/rebuilds those columns in MasterDocuments after a core refresh.
'
' Normal examples:
'   Field Name: Vessel Names
'   Search In: Document Title;Source File
'   Match Behavior: ALL UNIQUE
'
'   Keyword / Pattern: SAFEEN 3000
'   Result: SAFEEN-3000
'   Match Type: ALL TERMS
'
' Match types:
'   CONTAINS  - normal case-insensitive substring
'   ALL TERMS - spaces/+ are AND; -term excludes
'   EXACT     - exact case-insensitive whole search text
'   WILDCARD  - advanced ? fixed-width / * variable-width extraction concept

Public Sub NMDC_CustomFieldsInitialize()
    On Error GoTo Handler
    NMDC_LoadCustomFieldConfigurationIfEmpty
    NMDC_ApplyCustomFieldsUX
    Exit Sub
Handler:
    NMDC_LogError "CUSTOM_FIELDS_INIT_ERROR", _
        "Excel could not initialize Custom Fields & Keywords.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_OpenCustomFields()
    On Error GoTo Handler
    NMDC_LoadCustomFieldConfigurationIfEmpty
    NMDC_ApplyCustomFieldsUX
    NMDC_GoToSheet NMDC_CUSTOM_SHEET
    Exit Sub
Handler:
    NMDC_LogError "CUSTOM_FIELDS_OPEN_ERROR", _
        "Excel could not open Custom Fields & Keywords.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_AddCustomField()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim fields As ListObject
    Dim master As ListObject
    Dim newRow As ListRow
    Dim fieldName As String
    Dim searchIn As String
    Dim fieldColumn As ListColumn

    Set ws = ThisWorkbook.Worksheets(NMDC_CUSTOM_SHEET)
    Set fields = ws.ListObjects(NMDC_FIELDS_TABLE)
    Set master = ThisWorkbook.Worksheets("Master Documents").ListObjects("MasterDocuments")

    fieldName = Trim$(InputBox( _
        "Enter the new column name to add to Master Documents." & vbCrLf & vbCrLf & _
        "Example: Vessel Names", _
        "Add Custom Field"))
    If Len(fieldName) = 0 Then Exit Sub

    If NMDC_CustomFieldIsReserved(fieldName) Then
        MsgBox "That name belongs to a system/core field and cannot be used for a custom column." & vbCrLf & vbCrLf & _
               "Choose a new name such as Vessel Names, Asset Type, Contractor, Area, Campaign, etc.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    If NMDC_CustomFieldDefinitionExists(fields, fieldName) Then
        MsgBox "A Custom Fields definition already exists for: " & fieldName, _
               vbInformation, "NMDC Document Index"
        Exit Sub
    End If

    searchIn = Trim$(InputBox( _
        "Which Master Documents columns should be searched?" & vbCrLf & vbCrLf & _
        "Separate several columns with semicolons." & vbCrLf & _
        "Example: Document Title;Source File" & vbCrLf & _
        "Or type ALL TEXT to search all user-visible columns.", _
        "Search Columns", "Document Title;Source File"))
    If Len(searchIn) = 0 Then searchIn = "Document Title"

    Set newRow = NMDC_CustomBlankOrNewRow(fields)
    newRow.Range.Cells(1, fields.ListColumns("Enabled?").Index).Value = True
    newRow.Range.Cells(1, fields.ListColumns("Field Name").Index).Value = fieldName
    newRow.Range.Cells(1, fields.ListColumns("Search In").Index).Value = searchIn
    newRow.Range.Cells(1, fields.ListColumns("Match Behavior").Index).Value = "ALL UNIQUE"
    newRow.Range.Cells(1, fields.ListColumns("Separator").Index).Value = " | "
    newRow.Range.Cells(1, fields.ListColumns("Notes").Index).Value = "User-defined derived field"

    Set fieldColumn = Nothing
    On Error Resume Next
    Set fieldColumn = master.ListColumns(fieldName)
    On Error GoTo Handler
    If fieldColumn Is Nothing Then
        Set fieldColumn = master.ListColumns.Add
        fieldColumn.Name = fieldName
        If Not fieldColumn.DataBodyRange Is Nothing Then
            fieldColumn.DataBodyRange.ClearContents
            fieldColumn.DataBodyRange.NumberFormat = "@"
        End If
    End If

    NMDC_RebuildCustomFieldCheckboxes
    NMDC_ApplyCustomFieldsUX

    MsgBox "Custom field created: " & fieldName & vbCrLf & vbCrLf & _
           "Next, click Add Keyword Mapping and add the words/patterns that should populate this column.", _
           vbInformation, "NMDC Document Index"
    Exit Sub

Handler:
    NMDC_LogError "CUSTOM_FIELD_ADD_ERROR", _
        "Excel could not add the requested custom field.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_AddKeywordMapping()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim mappings As ListObject
    Dim fields As ListObject
    Dim newRow As ListRow
    Dim fieldName As String
    Dim keywordText As String
    Dim resultText As String
    Dim nextPriority As Long

    Set ws = ThisWorkbook.Worksheets(NMDC_CUSTOM_SHEET)
    Set fields = ws.ListObjects(NMDC_FIELDS_TABLE)
    Set mappings = ws.ListObjects(NMDC_MAPPINGS_TABLE)

    fieldName = NMDC_CustomSelectedFieldName(fields)
    If Len(fieldName) = 0 Then
        fieldName = Trim$(InputBox( _
            "Enter the Custom Field this keyword should populate." & vbCrLf & vbCrLf & _
            "Example: Vessel Names", _
            "Add Keyword Mapping"))
    End If
    If Len(fieldName) = 0 Then Exit Sub
    If Not NMDC_CustomFieldDefinitionExists(fields, fieldName) Then
        MsgBox "Create the Custom Field first, then add keyword mappings for it.", _
               vbExclamation, "NMDC Document Index"
        Exit Sub
    End If

    keywordText = Trim$(InputBox( _
        "Enter the keyword or pattern to find." & vbCrLf & vbCrLf & _
        "Examples:" & vbCrLf & _
        "SAFEEN-3000   (CONTAINS)" & vbCrLf & _
        "SAFEEN 3000   (ALL TERMS)" & vbCrLf & _
        "DLB-*         (WILDCARD - advanced)", _
        "Keyword / Pattern"))
    If Len(keywordText) = 0 Then Exit Sub

    resultText = Trim$(InputBox( _
        "Enter the value to write when this keyword matches." & vbCrLf & vbCrLf & _
        "Example: SAFEEN-3000", _
        "Result", keywordText))
    If Len(resultText) = 0 Then resultText = keywordText

    nextPriority = NMDC_CustomNextPriority(mappings, fieldName)
    Set newRow = NMDC_CustomBlankOrNewRow(mappings)
    newRow.Range.Cells(1, mappings.ListColumns("Enabled?").Index).Value = True
    newRow.Range.Cells(1, mappings.ListColumns("Field Name").Index).Value = fieldName
    newRow.Range.Cells(1, mappings.ListColumns("Keyword / Pattern").Index).Value = keywordText
    newRow.Range.Cells(1, mappings.ListColumns("Result").Index).Value = resultText
    newRow.Range.Cells(1, mappings.ListColumns("Match Type").Index).Value = "CONTAINS"
    newRow.Range.Cells(1, mappings.ListColumns("Priority").Index).Value = nextPriority
    newRow.Range.Cells(1, mappings.ListColumns("Notes").Index).Value = ""

    NMDC_RebuildCustomFieldCheckboxes
    NMDC_ApplyCustomFieldsUX
    Exit Sub

Handler:
    NMDC_LogError "KEYWORD_MAPPING_ADD_ERROR", _
        "Excel could not add the requested keyword mapping.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_ApplyCustomFields()
    If NMDC_ApplyCustomFieldsToMaster(True) Then
        NMDC_SaveCustomFieldConfiguration
    End If
End Sub

Public Function NMDC_ApplyCustomFieldsToMaster(Optional ByVal showMessage As Boolean = False) As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim fields As ListObject
    Dim mappings As ListObject
    Dim master As ListObject
    Dim fieldValues As Variant
    Dim mappingValues As Variant
    Dim masterValues As Variant
    Dim fieldRows As Long
    Dim mappingRows As Long
    Dim masterRows As Long
    Dim fieldRow As Long
    Dim rowIndex As Long
    Dim mapIndex As Long
    Dim orderCount As Long
    Dim fieldName As String
    Dim searchIn As String
    Dim behavior As String
    Dim separator As String
    Dim matchType As String
    Dim pattern As String
    Dim resultTemplate As String
    Dim matchedValue As String
    Dim finalValue As String
    Dim priority As Long
    Dim targetColumn As ListColumn
    Dim targetIndex As Long
    Dim sourceIndexes As Variant
    Dim mappingOrder() As Long
    Dim outputValues() As Variant
    Dim resultSet As Object
    Dim resultList As Collection
    Dim item As Variant
    Dim searchText As String
    Dim fieldsApplied As Long
    Dim mappedCells As Long
    Dim previousScreen As Boolean
    Dim previousEvents As Boolean
    Dim previousCalc As XlCalculation

    NMDC_LoadCustomFieldConfigurationIfEmpty

    Set ws = ThisWorkbook.Worksheets(NMDC_CUSTOM_SHEET)
    Set fields = ws.ListObjects(NMDC_FIELDS_TABLE)
    Set mappings = ws.ListObjects(NMDC_MAPPINGS_TABLE)
    Set master = ThisWorkbook.Worksheets("Master Documents").ListObjects("MasterDocuments")

    If master.DataBodyRange Is Nothing Then
        NMDC_ApplyCustomFieldsToMaster = True
        Exit Function
    End If
    If Not NMDC_CustomTableHasRows(fields, "Field Name") Then
        NMDC_ApplyCustomFieldsToMaster = True
        Exit Function
    End If

    fieldRows = fields.ListRows.Count
    mappingRows = mappings.ListRows.Count
    fieldValues = fields.DataBodyRange.Value2
    If Not mappings.DataBodyRange Is Nothing Then mappingValues = mappings.DataBodyRange.Value2

    previousScreen = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    previousCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' First ensure every enabled custom target column exists. Core fields are
    ' protected and can never be overwritten through this user layer.
    For fieldRow = 1 To fieldRows
        fieldName = Trim$(CStr(fieldValues(fieldRow, fields.ListColumns("Field Name").Index)))
        If Len(fieldName) > 0 Then
            If NMDC_CustomCheckedValue(fieldValues(fieldRow, fields.ListColumns("Enabled?").Index)) Then
                If NMDC_CustomFieldIsReserved(fieldName) Then
                    NMDC_LogError "CUSTOM_FIELD_RESERVED_NAME", _
                        "Custom field was skipped because its name conflicts with a core Master Documents field.", _
                        "Field: " & fieldName
                Else
                    Set targetColumn = Nothing
                    On Error Resume Next
                    Set targetColumn = master.ListColumns(fieldName)
                    On Error GoTo Handler
                    If targetColumn Is Nothing Then
                        Set targetColumn = master.ListColumns.Add
                        targetColumn.Name = fieldName
                    End If
                End If
            ElseIf Not NMDC_CustomFieldIsReserved(fieldName) Then
                Set targetColumn = Nothing
                On Error Resume Next
                Set targetColumn = master.ListColumns(fieldName)
                On Error GoTo Handler
                If Not targetColumn Is Nothing Then
                    If Not targetColumn.DataBodyRange Is Nothing Then targetColumn.DataBodyRange.ClearContents
                End If
            End If
        End If
    Next fieldRow

    masterRows = master.DataBodyRange.Rows.Count
    masterValues = master.DataBodyRange.Value2

    For fieldRow = 1 To fieldRows
        fieldName = Trim$(CStr(fieldValues(fieldRow, fields.ListColumns("Field Name").Index)))
        If Len(fieldName) = 0 Then GoTo NextField
        If Not NMDC_CustomCheckedValue(fieldValues(fieldRow, fields.ListColumns("Enabled?").Index)) Then GoTo NextField
        If NMDC_CustomFieldIsReserved(fieldName) Then GoTo NextField

        Set targetColumn = Nothing
        On Error Resume Next
        Set targetColumn = master.ListColumns(fieldName)
        On Error GoTo Handler
        If targetColumn Is Nothing Then GoTo NextField
        targetIndex = targetColumn.Index

        searchIn = Trim$(CStr(fieldValues(fieldRow, fields.ListColumns("Search In").Index)))
        If Len(searchIn) = 0 Then searchIn = "Document Title"
        behavior = UCase$(Trim$(CStr(fieldValues(fieldRow, fields.ListColumns("Match Behavior").Index))))
        If behavior <> "FIRST" Then behavior = "ALL UNIQUE"
        separator = CStr(fieldValues(fieldRow, fields.ListColumns("Separator").Index))
        If Len(separator) = 0 Then separator = " | "

        sourceIndexes = NMDC_CustomSourceIndexes(master, searchIn, targetIndex)
        orderCount = NMDC_CustomBuildMappingOrder( _
            mappings, mappingValues, mappingRows, fieldName, mappingOrder)

        ReDim outputValues(1 To masterRows, 1 To 1)

        For rowIndex = 1 To masterRows
            finalValue = ""
            searchText = NMDC_CustomComposeSearchText(masterValues, rowIndex, sourceIndexes)

            If orderCount > 0 And Len(searchText) > 0 Then
                If behavior = "FIRST" Then
                    For mapIndex = 1 To orderCount
                        pattern = Trim$(CStr(mappingValues(mappingOrder(mapIndex), mappings.ListColumns("Keyword / Pattern").Index)))
                        resultTemplate = CStr(mappingValues(mappingOrder(mapIndex), mappings.ListColumns("Result").Index))
                        matchType = UCase$(Trim$(CStr(mappingValues(mappingOrder(mapIndex), mappings.ListColumns("Match Type").Index))))
                        If Len(matchType) = 0 Then matchType = "CONTAINS"
                        matchedValue = ""
                        If NMDC_CustomEvaluateMatch(searchText, pattern, matchType, resultTemplate, matchedValue) Then
                            finalValue = matchedValue
                            Exit For
                        End If
                    Next mapIndex
                Else
                    Set resultSet = CreateObject("Scripting.Dictionary")
                    resultSet.CompareMode = vbTextCompare
                    Set resultList = New Collection

                    For mapIndex = 1 To orderCount
                        pattern = Trim$(CStr(mappingValues(mappingOrder(mapIndex), mappings.ListColumns("Keyword / Pattern").Index)))
                        resultTemplate = CStr(mappingValues(mappingOrder(mapIndex), mappings.ListColumns("Result").Index))
                        matchType = UCase$(Trim$(CStr(mappingValues(mappingOrder(mapIndex), mappings.ListColumns("Match Type").Index))))
                        If Len(matchType) = 0 Then matchType = "CONTAINS"
                        matchedValue = ""
                        If NMDC_CustomEvaluateMatch(searchText, pattern, matchType, resultTemplate, matchedValue) Then
                            If Len(matchedValue) > 0 Then
                                If Not resultSet.Exists(matchedValue) Then
                                    resultSet.Add matchedValue, True
                                    resultList.Add matchedValue
                                End If
                            End If
                        End If
                    Next mapIndex

                    For Each item In resultList
                        If Len(finalValue) > 0 Then finalValue = finalValue & separator
                        finalValue = finalValue & CStr(item)
                    Next item
                End If
            End If

            outputValues(rowIndex, 1) = finalValue
            masterValues(rowIndex, targetIndex) = finalValue
            If Len(finalValue) > 0 Then mappedCells = mappedCells + 1
        Next rowIndex

        targetColumn.DataBodyRange.NumberFormat = "@"
        targetColumn.DataBodyRange.Value2 = outputValues
        targetColumn.Range.ColumnWidth = 24
        fieldsApplied = fieldsApplied + 1

NextField:
    Next fieldRow

    NMDC_RebuildCustomFieldCheckboxes
    NMDC_ApplyCustomFieldsUX

CleanExit:
    Application.Calculation = previousCalc
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    NMDC_ApplyCustomFieldsToMaster = True

    If showMessage Then
        MsgBox "Custom fields updated." & vbCrLf & vbCrLf & _
               "Fields applied: " & CStr(fieldsApplied) & vbCrLf & _
               "Populated cells: " & CStr(mappedCells) & vbCrLf & vbCrLf & _
               "Core NMDC index fields were not changed.", _
               vbInformation, "NMDC Document Index"
    End If
    Exit Function

Handler:
    On Error Resume Next
    Application.Calculation = previousCalc
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    On Error GoTo 0
    NMDC_LogError "CUSTOM_FIELDS_APPLY_ERROR", _
        "Excel could not apply one or more user-defined custom fields.", _
        Err.Number & " - " & Err.Description
    NMDC_ApplyCustomFieldsToMaster = False
End Function

Public Sub NMDC_ApplyCustomFieldsUX()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim fields As ListObject
    Dim mappings As ListObject
    Dim area As Range
    Dim button As Shape

    Set ws = ThisWorkbook.Worksheets(NMDC_CUSTOM_SHEET)
    Set fields = ws.ListObjects(NMDC_FIELDS_TABLE)
    Set mappings = ws.ListObjects(NMDC_MAPPINGS_TABLE)

    On Error Resume Next
    ws.Range("A4:N4").UnMerge
    ws.Range("A4:N4").Merge
    On Error GoTo Handler
    With ws.Range("A4:N4")
        .Value = "CUSTOM FIELDS & KEYWORDS. Add a field (for example Vessel Names), choose which Master Documents columns to search, then add keyword mappings. Enabled rows use checkboxes. CONTAINS is simplest; ALL TERMS supports +AND / -EXCLUDE logic; WILDCARD (? and *) is advanced. Click Apply Custom Fields to rebuild the derived columns. Core index fields are protected."
        .Interior.Color = RGB(232, 241, 247)
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(31, 70, 90)
        .WrapText = True
        .RowHeight = 58
    End With

    NMDC_ApplyTableGuidance fields
    NMDC_ApplyTableGuidance mappings

    NMDC_CustomApplyListValidation fields, "Match Behavior", "FIRST,ALL UNIQUE"
    NMDC_CustomApplyListValidation mappings, "Match Type", "CONTAINS,ALL TERMS,EXACT,WILDCARD"

    On Error Resume Next
    ws.Shapes("NMDC_Custom_AddField").Delete
    ws.Shapes("NMDC_Custom_AddMap").Delete
    ws.Shapes("NMDC_Custom_Apply").Delete
    On Error GoTo Handler

    Set area = ws.Range("P2:R3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Custom_AddField"
    button.OnAction = "NMDC_AddCustomField"
    button.TextFrame.Characters.Text = "Add Custom Field" & vbLf & "Create Master Documents column"
    NMDC_CustomFormatButton button, RGB(20, 108, 148)

    Set area = ws.Range("S2:U3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Custom_AddMap"
    button.OnAction = "NMDC_AddKeywordMapping"
    button.TextFrame.Characters.Text = "Add Keyword Mapping" & vbLf & "Add a dictionary rule"
    NMDC_CustomFormatButton button, RGB(20, 108, 148)

    Set area = ws.Range("V2:X3")
    Set button = ws.Shapes.AddShape(5, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Custom_Apply"
    button.OnAction = "NMDC_ApplyCustomFields"
    button.TextFrame.Characters.Text = "Apply Custom Fields" & vbLf & "Save + rebuild columns"
    NMDC_CustomFormatButton button, RGB(46, 125, 50)

    NMDC_RebuildCustomFieldCheckboxes
    Exit Sub

Handler:
    NMDC_LogError "CUSTOM_FIELDS_UX_ERROR", _
        "Excel could not prepare the Custom Fields & Keywords editor.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_RebuildCustomFieldCheckboxes()
    On Error GoTo Handler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(NMDC_CUSTOM_SHEET)

    NMDC_CustomDeleteCheckboxes ws, NMDC_FIELD_CHECK_PREFIX
    NMDC_CustomDeleteCheckboxes ws, NMDC_MAP_CHECK_PREFIX

    NMDC_CustomBuildCheckboxes ws.ListObjects(NMDC_FIELDS_TABLE), "Enabled?", "Field Name", NMDC_FIELD_CHECK_PREFIX
    NMDC_CustomBuildCheckboxes ws.ListObjects(NMDC_MAPPINGS_TABLE), "Enabled?", "Keyword / Pattern", NMDC_MAP_CHECK_PREFIX
    Exit Sub

Handler:
    NMDC_LogError "CUSTOM_FIELD_CHECKBOX_BUILD_ERROR", _
        "Excel could not build all Custom Fields checkboxes.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_CustomCheckboxClicked()
    ' Compatibility stub for stale workbooks that still contain a legacy
    ' Form Control checkbox. Fresh builds use native in-cell checkboxes.
    On Error Resume Next
    NMDC_RebuildCustomFieldCheckboxes
    On Error GoTo 0
End Sub

Public Sub NMDC_SaveCustomFieldConfiguration()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim fields As ListObject
    Dim mappings As ListObject
    Dim root As String
    Dim fso As Object

    Set ws = ThisWorkbook.Worksheets(NMDC_CUSTOM_SHEET)
    Set fields = ws.ListObjects(NMDC_FIELDS_TABLE)
    Set mappings = ws.ListObjects(NMDC_MAPPINGS_TABLE)

    root = NMDC_RuntimePath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(root) Then fso.CreateFolder root

    NMDC_CustomWriteUtf8 fso.BuildPath(root, NMDC_CUSTOM_FIELDS_FILE), NMDC_CustomTableCsv(fields)
    NMDC_CustomWriteUtf8 fso.BuildPath(root, NMDC_KEYWORD_MAPPINGS_FILE), NMDC_CustomTableCsv(mappings)
    Exit Sub

Handler:
    NMDC_LogError "CUSTOM_FIELD_SAVE_ERROR", _
        "Excel could not save the Custom Fields configuration backup.", _
        Err.Number & " - " & Err.Description
End Sub

Public Sub NMDC_LoadCustomFieldConfigurationIfEmpty()
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim fields As ListObject
    Dim mappings As ListObject
    Dim fso As Object
    Dim fieldsPath As String
    Dim mappingsPath As String

    Set ws = ThisWorkbook.Worksheets(NMDC_CUSTOM_SHEET)
    Set fields = ws.ListObjects(NMDC_FIELDS_TABLE)
    Set mappings = ws.ListObjects(NMDC_MAPPINGS_TABLE)
    Set fso = CreateObject("Scripting.FileSystemObject")

    fieldsPath = fso.BuildPath(NMDC_RuntimePath(), NMDC_CUSTOM_FIELDS_FILE)
    mappingsPath = fso.BuildPath(NMDC_RuntimePath(), NMDC_KEYWORD_MAPPINGS_FILE)

    If Not NMDC_CustomTableHasRows(fields, "Field Name") Then
        If fso.FileExists(fieldsPath) Then NMDC_LoadCsvToTable fieldsPath, NMDC_CUSTOM_SHEET, NMDC_FIELDS_TABLE
    End If
    If Not NMDC_CustomTableHasRows(mappings, "Keyword / Pattern") Then
        If fso.FileExists(mappingsPath) Then NMDC_LoadCsvToTable mappingsPath, NMDC_CUSTOM_SHEET, NMDC_MAPPINGS_TABLE
    End If

    NMDC_RebuildCustomFieldCheckboxes
    Exit Sub

Handler:
    NMDC_LogError "CUSTOM_FIELD_LOAD_ERROR", _
        "Excel could not restore the saved Custom Fields configuration.", _
        Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_CustomBuildMappingOrder( _
    ByVal mappings As ListObject, ByVal values As Variant, ByVal rowCount As Long, _
    ByVal fieldName As String, ByRef order() As Long) As Long

    Dim rowIndex As Long
    Dim count As Long
    Dim insertAt As Long
    Dim keyRow As Long
    Dim keyPriority As Long
    Dim fieldCol As Long
    Dim enabledCol As Long
    Dim priorityCol As Long

    If mappings.DataBodyRange Is Nothing Then Exit Function

    fieldCol = mappings.ListColumns("Field Name").Index
    enabledCol = mappings.ListColumns("Enabled?").Index
    priorityCol = mappings.ListColumns("Priority").Index

    ReDim order(1 To rowCount)
    count = 0
    For rowIndex = 1 To rowCount
        If StrComp(Trim$(CStr(values(rowIndex, fieldCol))), fieldName, vbTextCompare) = 0 Then
            If NMDC_CustomCheckedValue(values(rowIndex, enabledCol)) Then
                If Len(Trim$(CStr(values(rowIndex, mappings.ListColumns("Keyword / Pattern").Index)))) > 0 Then
                    count = count + 1
                    order(count) = rowIndex
                End If
            End If
        End If
    Next rowIndex

    ' Stable insertion sort by numeric priority (lower runs first).
    For rowIndex = 2 To count
        keyRow = order(rowIndex)
        keyPriority = NMDC_CustomPriority(values(keyRow, priorityCol))
        insertAt = rowIndex - 1
        Do While insertAt >= 1
            If NMDC_CustomPriority(values(order(insertAt), priorityCol)) <= keyPriority Then Exit Do
            order(insertAt + 1) = order(insertAt)
            insertAt = insertAt - 1
        Loop
        order(insertAt + 1) = keyRow
    Next rowIndex

    NMDC_CustomBuildMappingOrder = count
End Function

Private Function NMDC_CustomSourceIndexes(ByVal master As ListObject, ByVal searchIn As String, ByVal targetIndex As Long) As Variant
    Dim indexes() As Long
    Dim count As Long
    Dim columnIndex As Long
    Dim parts As Variant
    Dim part As Variant
    Dim nameText As String
    Dim found As Long

    count = 0
    searchIn = Trim$(searchIn)

    If UCase$(searchIn) = "ALL TEXT" Or UCase$(searchIn) = "*ALL TEXT*" Then
        ReDim indexes(1 To master.ListColumns.Count)
        For columnIndex = 1 To master.ListColumns.Count
            If columnIndex <> targetIndex Then
                If Left$(CStr(master.ListColumns(columnIndex).Name), 7) <> "__NMDC_" Then
                    count = count + 1
                    indexes(count) = columnIndex
                End If
            End If
        Next columnIndex
    Else
        parts = Split(Replace(searchIn, ",", ";"), ";")
        ReDim indexes(1 To UBound(parts) + 1)
        For Each part In parts
            nameText = Trim$(CStr(part))
            If Len(nameText) > 0 Then
                found = 0
                On Error Resume Next
                found = master.ListColumns(nameText).Index
                On Error GoTo 0
                If found > 0 And found <> targetIndex Then
                    If Left$(CStr(master.ListColumns(found).Name), 7) <> "__NMDC_" Then
                        count = count + 1
                        indexes(count) = found
                    End If
                End If
            End If
        Next part
    End If

    If count = 0 Then
        ReDim indexes(1 To 1)
        On Error Resume Next
        indexes(1) = master.ListColumns("Document Title").Index
        On Error GoTo 0
        If indexes(1) = 0 Then indexes(1) = 1
        count = 1
    End If

    ReDim Preserve indexes(1 To count)
    NMDC_CustomSourceIndexes = indexes
End Function

Private Function NMDC_CustomComposeSearchText(ByVal values As Variant, ByVal rowIndex As Long, ByVal indexes As Variant) As String
    Dim i As Long
    Dim partText As String
    Dim combined As String

    combined = ""
    For i = LBound(indexes) To UBound(indexes)
        partText = CStr(values(rowIndex, CLng(indexes(i))))
        If Len(partText) > 0 Then
            If Len(combined) > 0 Then combined = combined & " | "
            combined = combined & partText
        End If
    Next i
    NMDC_CustomComposeSearchText = combined
End Function

Private Function NMDC_CustomEvaluateMatch( _
    ByVal searchText As String, ByVal pattern As String, ByVal matchType As String, _
    ByVal resultTemplate As String, ByRef outputValue As String) As Boolean

    Dim matched As Boolean
    Dim extracted As String

    outputValue = ""
    pattern = Trim$(pattern)
    If Len(pattern) = 0 Then Exit Function

    Select Case UCase$(Trim$(matchType))
        Case "ALL TERMS"
            matched = NMDC_CustomAllTermsMatch(searchText, pattern)
        Case "EXACT"
            matched = (StrComp(Trim$(searchText), pattern, vbTextCompare) = 0)
        Case "WILDCARD"
            matched = NMDC_CustomWildcardMatch(searchText, pattern, extracted)
        Case Else
            matched = (InStr(1, searchText, pattern, vbTextCompare) > 0)
    End Select

    If Not matched Then Exit Function

    If Len(resultTemplate) = 0 Then
        If Len(extracted) > 0 Then
            outputValue = extracted
        Else
            outputValue = pattern
        End If
    ElseIf Len(extracted) > 0 And InStr(1, resultTemplate, "*", vbBinaryCompare) > 0 Then
        outputValue = Replace(resultTemplate, "*", extracted)
    ElseIf Len(extracted) > 0 And InStr(1, resultTemplate, String$(Len(extracted), "?"), vbBinaryCompare) > 0 Then
        outputValue = Replace(resultTemplate, String$(Len(extracted), "?"), extracted)
    Else
        outputValue = resultTemplate
    End If

    outputValue = Trim$(outputValue)
    NMDC_CustomEvaluateMatch = True
End Function

Private Function NMDC_CustomAllTermsMatch(ByVal searchText As String, ByVal expressionText As String) As Boolean
    Dim tokens As Variant
    Dim token As Variant
    Dim cleanToken As String
    Dim positiveCount As Long

    NMDC_CustomAllTermsMatch = True
    tokens = Split(Replace(Trim$(expressionText), "+", " "), " ")

    For Each token In tokens
        cleanToken = Trim$(CStr(token))
        If Len(cleanToken) > 0 Then
            If Left$(cleanToken, 1) = "-" And Len(cleanToken) > 1 Then
                If InStr(1, searchText, Mid$(cleanToken, 2), vbTextCompare) > 0 Then
                    NMDC_CustomAllTermsMatch = False
                    Exit Function
                End If
            Else
                positiveCount = positiveCount + 1
                If InStr(1, searchText, cleanToken, vbTextCompare) = 0 Then
                    NMDC_CustomAllTermsMatch = False
                    Exit Function
                End If
            End If
        End If
    Next token
End Function

Private Function NMDC_CustomWildcardMatch(ByVal searchText As String, ByVal pattern As String, ByRef extracted As String) As Boolean
    On Error GoTo Failed

    Dim starCount As Long
    Dim questionCount As Long
    Dim parts As Variant
    Dim prefix As String
    Dim suffix As String
    Dim startPos As Long
    Dim endPos As Long
    Dim cleanPattern As String
    Dim leadingQuestions As Long
    Dim trailingQuestions As Long
    Dim upperText As String
    Dim upperPattern As String

    extracted = ""
    starCount = Len(pattern) - Len(Replace(pattern, "*", ""))
    questionCount = Len(pattern) - Len(Replace(pattern, "?", ""))

    If starCount = 1 Then
        parts = Split(pattern, "*")
        prefix = CStr(parts(0))
        suffix = CStr(parts(1))

        If Len(prefix) > 0 Then
            startPos = InStr(1, searchText, prefix, vbTextCompare)
            If startPos = 0 Then Exit Function
            startPos = startPos + Len(prefix)
        Else
            startPos = 1
        End If

        If Len(suffix) > 0 Then
            endPos = InStr(startPos, searchText, suffix, vbTextCompare)
            If endPos = 0 Then Exit Function
        Else
            endPos = Len(searchText) + 1
        End If

        If endPos < startPos Then Exit Function
        extracted = Mid$(searchText, startPos, endPos - startPos)
        NMDC_CustomWildcardMatch = True
        Exit Function
    End If

    If questionCount > 0 Then
        leadingQuestions = NMDC_CustomLeadingCharCount(pattern, "?")
        trailingQuestions = NMDC_CustomTrailingCharCount(pattern, "?")
        cleanPattern = Replace(pattern, "?", "")

        If trailingQuestions = questionCount And Len(cleanPattern) > 0 Then
            startPos = InStr(1, searchText, cleanPattern, vbTextCompare)
            If startPos > 0 Then
                startPos = startPos + Len(cleanPattern)
                If startPos + questionCount - 1 <= Len(searchText) Then
                    extracted = Mid$(searchText, startPos, questionCount)
                    NMDC_CustomWildcardMatch = True
                    Exit Function
                End If
            End If
        ElseIf leadingQuestions = questionCount And Len(cleanPattern) > 0 Then
            endPos = InStr(1, searchText, cleanPattern, vbTextCompare)
            If endPos > questionCount Then
                extracted = Mid$(searchText, endPos - questionCount, questionCount)
                NMDC_CustomWildcardMatch = True
                Exit Function
            End If
        End If
    End If

    ' General wildcard fallback. This provides matching even when the pattern is
    ' more complex than the extraction forms above; static Result values still work.
    upperText = UCase$(searchText)
    upperPattern = UCase$(pattern)
    If upperText Like "*" & upperPattern & "*" Then
        NMDC_CustomWildcardMatch = True
        Exit Function
    End If

Failed:
    NMDC_CustomWildcardMatch = False
End Function

Private Function NMDC_CustomLeadingCharCount(ByVal text As String, ByVal oneChar As String) As Long
    Dim i As Long
    For i = 1 To Len(text)
        If Mid$(text, i, 1) <> oneChar Then Exit For
        NMDC_CustomLeadingCharCount = NMDC_CustomLeadingCharCount + 1
    Next i
End Function

Private Function NMDC_CustomTrailingCharCount(ByVal text As String, ByVal oneChar As String) As Long
    Dim i As Long
    For i = Len(text) To 1 Step -1
        If Mid$(text, i, 1) <> oneChar Then Exit For
        NMDC_CustomTrailingCharCount = NMDC_CustomTrailingCharCount + 1
    Next i
End Function

Private Function NMDC_CustomFieldIsReserved(ByVal fieldName As String) As Boolean
    If Left$(UCase$(Trim$(fieldName)), 7) = "__NMDC_" Then
        NMDC_CustomFieldIsReserved = True
        Exit Function
    End If

    Select Case UCase$(Trim$(fieldName))
        Case "FLAG LEVEL", "PROJECT NO.", "SOURCE FAMILY", "DISCIPLINE", "CATEGORY", "SUBCATEGORY", _
             "DOCUMENT NO.", "DOCUMENT TITLE", "COMPANY DOCUMENT NO.", "LATEST REVISION", _
             "LATEST EVENT DATE", "LATEST EVENT STATUS", "DOCUMENT LINK", "SOURCE FILE", _
             "SOURCE SHEET", "SOURCE ROW", "SOURCE CELL", "GLOBAL DOCUMENT KEY"
            NMDC_CustomFieldIsReserved = True
        Case Else
            NMDC_CustomFieldIsReserved = False
    End Select
End Function

Private Function NMDC_CustomFieldDefinitionExists(ByVal fields As ListObject, ByVal fieldName As String) As Boolean
    Dim row As ListRow
    Dim fieldCol As Long
    fieldCol = fields.ListColumns("Field Name").Index
    For Each row In fields.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, fieldCol).Value)), fieldName, vbTextCompare) = 0 Then
            NMDC_CustomFieldDefinitionExists = True
            Exit Function
        End If
    Next row
End Function

Private Function NMDC_CustomSelectedFieldName(ByVal fields As ListObject) As String
    On Error GoTo Failed
    Dim hit As Range
    Dim rowIndex As Long
    If TypeName(Selection) <> "Range" Then Exit Function
    If Selection.Worksheet.Name <> NMDC_CUSTOM_SHEET Then Exit Function
    Set hit = Intersect(Selection.Cells(1, 1), fields.DataBodyRange)
    If hit Is Nothing Then Exit Function
    rowIndex = Selection.Cells(1, 1).Row - fields.DataBodyRange.Row + 1
    NMDC_CustomSelectedFieldName = Trim$(CStr(fields.ListColumns("Field Name").DataBodyRange.Cells(rowIndex, 1).Value))
    Exit Function
Failed:
    NMDC_CustomSelectedFieldName = ""
End Function

Private Function NMDC_CustomNextPriority(ByVal mappings As ListObject, ByVal fieldName As String) As Long
    Dim row As ListRow
    Dim currentPriority As Long
    Dim maxPriority As Long
    Dim fieldCol As Long
    Dim priorityCol As Long

    fieldCol = mappings.ListColumns("Field Name").Index
    priorityCol = mappings.ListColumns("Priority").Index
    For Each row In mappings.ListRows
        If StrComp(Trim$(CStr(row.Range.Cells(1, fieldCol).Value)), fieldName, vbTextCompare) = 0 Then
            currentPriority = NMDC_CustomPriority(row.Range.Cells(1, priorityCol).Value)
            If currentPriority < 999999 And currentPriority > maxPriority Then maxPriority = currentPriority
        End If
    Next row
    NMDC_CustomNextPriority = maxPriority + 10
    If NMDC_CustomNextPriority < 10 Then NMDC_CustomNextPriority = 10
End Function

Private Function NMDC_CustomPriority(ByVal value As Variant) As Long
    If IsNumeric(value) Then
        NMDC_CustomPriority = CLng(value)
    Else
        NMDC_CustomPriority = 999999
    End If
End Function

Private Function NMDC_CustomCheckedValue(ByVal value As Variant) As Boolean
    If VarType(value) = vbBoolean Then
        NMDC_CustomCheckedValue = CBool(value)
    Else
        Select Case UCase$(Trim$(CStr(value)))
            Case "TRUE", "YES", "Y", "1", "ON", "CHECKED"
                NMDC_CustomCheckedValue = True
            Case Else
                NMDC_CustomCheckedValue = False
        End Select
    End If
End Function

Private Function NMDC_CustomBlankOrNewRow(ByVal table As ListObject) As ListRow
    If table.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(table.ListRows(1).Range) = 0 Then
            Set NMDC_CustomBlankOrNewRow = table.ListRows(1)
            Exit Function
        End If
    End If
    Set NMDC_CustomBlankOrNewRow = table.ListRows.Add
End Function

Private Function NMDC_CustomTableHasRows(ByVal table As ListObject, ByVal keyColumn As String) As Boolean
    Dim cell As Range
    If table.DataBodyRange Is Nothing Then Exit Function
    For Each cell In table.ListColumns(keyColumn).DataBodyRange.Cells
        If Len(Trim$(CStr(cell.Value))) > 0 Then
            NMDC_CustomTableHasRows = True
            Exit Function
        End If
    Next cell
End Function

Private Sub NMDC_CustomDeleteCheckboxes(ByVal ws As Worksheet, ByVal prefix As String)
    Dim i As Long
    Dim item As Object
    For i = ws.CheckBoxes.Count To 1 Step -1
        Set item = ws.CheckBoxes(i)
        If Left$(CStr(item.Name), Len(prefix)) = prefix Then item.Delete
    Next i
End Sub

Private Sub NMDC_CustomBuildCheckboxes( _
    ByVal table As ListObject, ByVal enabledColumnName As String, ByVal keyColumnName As String, ByVal prefix As String)

    Dim rowIndex As Long
    Dim targetCell As Range
    Dim keyText As String
    Dim isIncluded As Boolean

    If table.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To table.ListRows.Count
        keyText = Trim$(CStr(table.ListColumns(keyColumnName).DataBodyRange.Cells(rowIndex, 1).Value))
        Set targetCell = table.ListColumns(enabledColumnName).DataBodyRange.Cells(rowIndex, 1)
        If Len(keyText) > 0 Then
            isIncluded = NMDC_CustomCheckedValue(targetCell.Value)
            targetCell.Value = isIncluded
            targetCell.NumberFormat = "General"
            targetCell.HorizontalAlignment = xlCenter
            On Error Resume Next
            targetCell.CellControl.SetCheckbox
            On Error GoTo 0
        Else
            targetCell.ClearContents
        End If
    Next rowIndex
End Sub

Private Sub NMDC_CustomApplyListValidation(ByVal table As ListObject, ByVal columnName As String, ByVal listValues As String)
    Dim target As Range
    On Error Resume Next
    Set target = table.ListColumns(columnName).DataBodyRange
    On Error GoTo 0
    If target Is Nothing Then Exit Sub
    target.Validation.Delete
    target.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, listValues
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
End Sub

Private Sub NMDC_CustomFormatButton(ByVal button As Shape, ByVal fillColor As Long)
    button.TextFrame.HorizontalAlignment = xlCenter
    button.TextFrame.VerticalAlignment = xlCenter
    button.Fill.ForeColor.RGB = fillColor
    button.Line.ForeColor.RGB = fillColor
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 9
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    button.Placement = xlMoveAndSize
End Sub

Private Function NMDC_CustomTableCsv(ByVal table As ListObject) As String
    Dim text As String
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim value As Variant

    For columnIndex = 1 To table.ListColumns.Count
        If columnIndex > 1 Then text = text & ","
        text = text & NMDC_CustomCsvField(CStr(table.ListColumns(columnIndex).Name))
    Next columnIndex
    text = text & vbLf

    If Not table.DataBodyRange Is Nothing Then
        For rowIndex = 1 To table.ListRows.Count
            If Application.WorksheetFunction.CountA(table.ListRows(rowIndex).Range) > 0 Then
                For columnIndex = 1 To table.ListColumns.Count
                    If columnIndex > 1 Then text = text & ","
                    value = table.DataBodyRange.Cells(rowIndex, columnIndex).Value
                    If VarType(value) = vbBoolean Then
                        text = text & NMDC_CustomCsvField(IIf(CBool(value), "TRUE", "FALSE"))
                    Else
                        text = text & NMDC_CustomCsvField(CStr(value))
                    End If
                Next columnIndex
                text = text & vbLf
            End If
        Next rowIndex
    End If

    NMDC_CustomTableCsv = text
End Function

Private Function NMDC_CustomCsvField(ByVal value As String) As String
    value = Replace(value, Chr$(34), Chr$(34) & Chr$(34))
    value = Replace(value, vbCrLf, " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    NMDC_CustomCsvField = Chr$(34) & value & Chr$(34)
End Function

Private Sub NMDC_CustomWriteUtf8(ByVal filePath As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close
End Sub
