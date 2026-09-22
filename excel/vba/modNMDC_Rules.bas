Attribute VB_Name = "modNMDC_Rules"
Option Explicit

Private Const RULES_FILENAME As String = "classification_rules.csv"

Public Function NMDC_PrepareRulesForStage() As Boolean
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim csvText As String
    Dim targetPath As String
    Dim existingText As String

    Set ws = ThisWorkbook.Worksheets("Rules & Mappings")
    Set table = ws.ListObjects("ClassificationRules")
    If Not NMDC_ValidateRules(table) Then
        NMDC_PrepareRulesForStage = False
        Exit Function
    End If

    csvText = NMDC_RulesCsv(table)
    NMDC_EnsureFolder NMDC_ConfigPath()
    targetPath = NMDC_ConfigPath() & "\" & RULES_FILENAME
    existingText = NMDC_ReadUtf8(targetPath)
    If StrComp(existingText, csvText, vbBinaryCompare) <> 0 Then
        NMDC_WriteUtf8 targetPath, csvText
        NMDC_SetConfigValue "Configuration Version", Format$(Now, "yyyymmdd-hhnnss")
        NMDC_RecordRuleChange "Classification rules updated from Excel before staging."
        MsgBox "Rules & Mappings were validated and saved." & vbCrLf & vbCrLf & _
               "The update will reprocess the affected source data and remain staged until you approve it.", _
               vbInformation, "NMDC Document Index"
    End If
    NMDC_PrepareRulesForStage = True
    Exit Function

Handler:
    NMDC_LogError "RULE_EXPORT_ERROR", _
        "Excel could not validate or save Rules & Mappings.", _
        Err.Number & " - " & Err.Description & "; Rules target=" & NMDC_ConfigPath() & "\" & RULES_FILENAME & "; " & NMDC_PathDiagnostics()
    MsgBox "Rules & Mappings could not be saved." & vbCrLf & _
           "Nothing was staged. Please review the Error Log.", _
           vbExclamation, "NMDC Document Index"
    NMDC_PrepareRulesForStage = False
End Function

Private Function NMDC_ValidateRules(ByVal table As ListObject) As Boolean
    Dim required As Variant
    Dim item As Variant
    Dim seen As Object
    Dim row As ListRow
    Dim ruleId As String
    Dim enabled As String
    Dim includeValue As String

    required = Array("Rule_ID", "Enabled", "Priority", "Source_Family", "Match_Scope", _
                     "Match_Type", "Match_Words", "Discipline", "Category", "Subcategory", "Include")
    For Each item In required
        If NMDC_TableColumn(table, CStr(item)) = 0 Then
            MsgBox "Rules & Mappings is missing the required column: " & CStr(item), _
                   vbExclamation, "NMDC Document Index"
            Exit Function
        End If
    Next item

    Set seen = CreateObject("Scripting.Dictionary")
    For Each row In table.ListRows
        ruleId = Trim$(CStr(row.Range.Cells(1, NMDC_TableColumn(table, "Rule_ID")).Value))
        If Len(ruleId) > 0 Then
            If seen.Exists(UCase$(ruleId)) Then
                MsgBox "Duplicate Rule_ID: " & ruleId & vbCrLf & _
                       "Nothing was staged.", vbExclamation, "NMDC Document Index"
                Exit Function
            End If
            seen.Add UCase$(ruleId), True

            enabled = UCase$(Trim$(CStr(row.Range.Cells(1, NMDC_TableColumn(table, "Enabled")).Value)))
            includeValue = UCase$(Trim$(CStr(row.Range.Cells(1, NMDC_TableColumn(table, "Include")).Value)))
            If enabled <> "YES" And enabled <> "NO" Then
                MsgBox "Rule " & ruleId & " has an invalid Enabled value. Use YES or NO.", _
                       vbExclamation, "NMDC Document Index"
                Exit Function
            End If
            If includeValue <> "YES" And includeValue <> "NO" Then
                MsgBox "Rule " & ruleId & " has an invalid Include value. Use YES or NO.", _
                       vbExclamation, "NMDC Document Index"
                Exit Function
            End If
            If Len(Trim$(CStr(row.Range.Cells(1, NMDC_TableColumn(table, "Match_Words")).Value))) = 0 Then
                MsgBox "Rule " & ruleId & " has no Match_Words value.", _
                       vbExclamation, "NMDC Document Index"
                Exit Function
            End If
        End If
    Next row
    NMDC_ValidateRules = True
End Function

Private Function NMDC_TableColumn(ByVal table As ListObject, ByVal headerName As String) As Long
    Dim column As ListColumn
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), headerName, vbTextCompare) = 0 Then
            NMDC_TableColumn = column.Index
            Exit Function
        End If
    Next column
    NMDC_TableColumn = 0
End Function

Private Function NMDC_RulesCsv(ByVal table As ListObject) As String
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim lineText As String
    Dim outputText As String
    Dim ruleIdColumn As Long

    ruleIdColumn = NMDC_TableColumn(table, "Rule_ID")
    For columnIndex = 1 To table.ListColumns.Count
        If columnIndex > 1 Then lineText = lineText & ","
        lineText = lineText & NMDC_CsvField(CStr(table.HeaderRowRange.Cells(1, columnIndex).Value))
    Next columnIndex
    outputText = lineText & vbLf

    For rowIndex = 1 To table.ListRows.Count
        If ruleIdColumn = 0 Or Len(Trim$(CStr(table.DataBodyRange.Cells(rowIndex, ruleIdColumn).Value))) > 0 Then
            lineText = ""
            For columnIndex = 1 To table.ListColumns.Count
                If columnIndex > 1 Then lineText = lineText & ","
                lineText = lineText & NMDC_CsvField(CStr(table.DataBodyRange.Cells(rowIndex, columnIndex).Value))
            Next columnIndex
            outputText = outputText & lineText & vbLf
        End If
    Next rowIndex
    NMDC_RulesCsv = outputText
End Function

Private Function NMDC_CsvField(ByVal value As String) As String
    NMDC_CsvField = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Sub NMDC_EnsureFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Function NMDC_ReadUtf8(ByVal filePath As String) As String
    On Error GoTo Missing
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    NMDC_ReadUtf8 = stream.ReadText
    stream.Close
    Exit Function
Missing:
    NMDC_ReadUtf8 = ""
End Function

Private Sub NMDC_WriteUtf8(ByVal filePath As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText value
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

Private Sub NMDC_RecordRuleChange(ByVal note As String)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim table As ListObject
    Dim row As ListRow

    Set ws = ThisWorkbook.Worksheets("User Decisions")
    Set table = ws.ListObjects("UserDecisionLog")
    Set row = NMDC_BlankOrNewDecisionRow(table)

    NMDC_SetDecisionValue table, row, "Decision ID", "RULE-" & Format$(Now, "yyyymmddhhnnss")
    NMDC_SetDecisionValue table, row, "Date/Time", Now
    NMDC_SetDecisionValue table, row, "User", Application.UserName
    NMDC_SetDecisionValue table, row, "Decision Type", "CONFIGURATION CHANGE"
    NMDC_SetDecisionValue table, row, "User Note", note
    NMDC_SetDecisionValue table, row, "Status", "STAGED ON NEXT UPDATE"
    NMDC_SetDecisionValue table, row, "Configuration Version", NMDC_ConfigValue("Configuration Version")
    table.ListColumns("Date/Time").DataBodyRange.NumberFormat = "dd-mmm-yyyy hh:mm"
    Exit Sub
Handler:
    NMDC_LogError "RULE_CHANGE_LOG_ERROR", _
        "The rules were saved, but Excel could not record the configuration change in User Decisions.", _
        Err.Number & " - " & Err.Description
End Sub

Private Function NMDC_BlankOrNewDecisionRow(ByVal table As ListObject) As ListRow
    If table.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(table.ListRows(1).Range) = 0 Then
            Set NMDC_BlankOrNewDecisionRow = table.ListRows(1)
            Exit Function
        End If
    End If
    Set NMDC_BlankOrNewDecisionRow = table.ListRows.Add
End Function

Private Sub NMDC_SetDecisionValue(ByVal table As ListObject, ByVal row As ListRow, ByVal headerName As String, ByVal value As Variant)
    row.Range.Cells(1, table.ListColumns(headerName).Index).Value = value
End Sub
