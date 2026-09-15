Option Explicit

Const xlOpenXMLWorkbookMacroEnabled = 52
Const xlSrcRange = 1
Const xlYes = 1
Const xlUp = -4162
Const xlToLeft = -4159
Const xlValidateList = 3
Const xlValidAlertStop = 1
Const xlBetween = 1
Const msoShapeRoundedRectangle = 5

Dim fso, shell, packageRoot, sourceWorkbook, outputWorkbook, excel, workbook
Dim modulesFolder, enginePath, runtimeFolder, configFolder, response

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
packageRoot = fso.GetParentFolderName(WScript.ScriptFullName)
sourceWorkbook = fso.BuildPath(packageRoot, "source\NMDC_Document_Index_Base.xlsx")
outputWorkbook = fso.BuildPath(packageRoot, "NMDC_Document_Index.xlsm")
modulesFolder = fso.BuildPath(packageRoot, "vba")
enginePath = fso.BuildPath(packageRoot, "engine\nmdc_index_engine.exe")
runtimeFolder = fso.BuildPath(packageRoot, "runtime")
configFolder = fso.BuildPath(packageRoot, "config")

If Not fso.FileExists(sourceWorkbook) Then
    MsgBox "The production workbook source is missing:" & vbCrLf & sourceWorkbook, vbCritical, "NMDC Document Index Setup"
    WScript.Quit 2
End If
If Not fso.FolderExists(modulesFolder) Then
    MsgBox "The VBA source folder is missing:" & vbCrLf & modulesFolder, vbCritical, "NMDC Document Index Setup"
    WScript.Quit 2
End If
RequireFile enginePath, "The packaged NMDC Index engine is missing."
RequireFile fso.BuildPath(configFolder, "classification_rules.csv"), "The classification rules file is missing."
RequireFile fso.BuildPath(configFolder, "project_identity_overrides.csv"), "The project identity overrides file is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Engine.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Csv.bas"), "A required safe CSV module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Refresh.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Actions.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_TableActions.bas"), "A required Excel table-action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Admin.bas"), "A required Excel administration module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Rules.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Startup.bas"), "A required Excel action module is missing."
If Not fso.FolderExists(runtimeFolder) Then fso.CreateFolder runtimeFolder

If fso.FileExists(outputWorkbook) Then
    response = MsgBox("NMDC_Document_Index.xlsm already exists." & vbCrLf & vbCrLf & _
                      "Create a backup and replace it?", vbQuestion + vbYesNo + vbDefaultButton2, "NMDC Document Index Setup")
    If response <> vbYes Then WScript.Quit 1
    fso.CopyFile outputWorkbook, outputWorkbook & ".backup-" & SafeTimestamp(), True
End If

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Microsoft Excel desktop could not be started. Please confirm that Excel is installed.", vbCritical, "NMDC Document Index Setup"
    WScript.Quit 3
End If
Err.Clear

excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Open(sourceWorkbook, False, False)
If Err.Number <> 0 Then
    ShowFailure "Excel could not open the production workbook source."
    WScript.Quit 4
End If

Dim projectCheck
Set projectCheck = workbook.VBProject
If Err.Number <> 0 Then
    workbook.Close False
    excel.Quit
    MsgBox "Excel blocked the one-time production setup." & vbCrLf & vbCrLf & _
           "In Excel, open File > Options > Trust Center > Trust Center Settings > Macro Settings." & vbCrLf & _
           "Temporarily select 'Trust access to the VBA project object model', then run this setup again." & vbCrLf & vbCrLf & _
           "After the workbook is created, you may turn that setting off again.", _
           vbExclamation, "NMDC Document Index Setup"
    WScript.Quit 5
End If
Err.Clear

workbook.SaveAs outputWorkbook, xlOpenXMLWorkbookMacroEnabled
If Err.Number <> 0 Then
    ShowFailure "Excel could not create NMDC_Document_Index.xlsm."
    WScript.Quit 6
End If

EnsureNamedTables workbook
If Err.Number <> 0 Then
    ShowFailure "Excel could not validate the required named tables."
    WScript.Quit 7
End If

ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Engine.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Csv.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Refresh.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Actions.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_TableActions.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Admin.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Rules.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Startup.bas")
If Err.Number <> 0 Then
    ShowFailure "Excel could not attach the production actions."
    WScript.Quit 8
End If

SetWorkbookConfig workbook, "Engine Executable Path", enginePath
SetWorkbookConfig workbook, "Runtime Folder", runtimeFolder
SetWorkbookConfig workbook, "Configuration Folder", configFolder
SetWorkbookConfig workbook, "Classification Rules File", fso.BuildPath(configFolder, "classification_rules.csv")
SetWorkbookConfig workbook, "Project Identity Overrides File", fso.BuildPath(configFolder, "project_identity_overrides.csv")

StyleHomeDashboard workbook
AttachHomeButtons workbook
ConfigureReviewFlags workbook
ApplyUserDropdowns workbook
workbook.Worksheets("System Data").Visible = 2
NormalizeMergedUiRanges workbook
workbook.Save
workbook.Close True
excel.Quit
On Error GoTo 0

MsgBox "NMDC_Document_Index.xlsm was created successfully." & vbCrLf & vbCrLf & _
       "Open it in Microsoft Excel, enable macros, select the DATA folder, and press Full Rescan / Rebuild All for the owner retest.", _
       vbInformation, "NMDC Document Index Setup"

Sub ImportModule(ByVal wb, ByVal modulePath)
    If Not fso.FileExists(modulePath) Then
        Err.Raise vbObjectError + 100, "NMDC Setup", "Missing VBA module: " & modulePath
    End If
    wb.VBProject.VBComponents.Import modulePath
End Sub

Sub RequireFile(ByVal filePath, ByVal friendlyMessage)
    If Not fso.FileExists(filePath) Then
        MsgBox friendlyMessage & vbCrLf & vbCrLf & filePath & vbCrLf & vbCrLf & _
               "Extract the complete production ZIP to a normal folder and run setup again.", _
               vbCritical, "NMDC Document Index Setup"
        WScript.Quit 2
    End If
End Sub

Sub EnsureNamedTables(ByVal wb)
    EnsureTable wb, "Master Documents", "MasterDocuments", 5, Empty
    EnsureTable wb, "Revisions", "RevisionRegister", 5, Empty
    EnsureTable wb, "Transactions", "EventRegister", 5, Empty
    EnsureTable wb, "Pending Update", "PendingUpdate", 5, Array("Change Type", "Record Identity", "Project No.", "Document No.", "Revision", "Event Type", "Source File", "Plain-English Summary", "Review Required")
    EnsureTable wb, "Review Flags", "ReviewFlags", 5, Empty
    EnsureTable wb, "User Decisions", "UserDecisionLog", 11, Empty
    EnsureTable wb, "Configuration", "Configuration", 5, Empty
    EnsureTable wb, "Rules & Mappings", "ClassificationRules", 5, Empty
    EnsureTable wb, "Update History", "UpdateHistory", 5, Empty
    EnsureTable wb, "Error Log", "ErrorLog", 5, Array("Date/Time", "Severity", "Action", "Plain-English Error", "Recommended Action", "Technical Detail", "Run ID", "Source File", "Worksheet", "Source Row/Cell")
    EnsureTable wb, "System Data", "BaselineCounts", 5, Empty
    EnsureTable wb, "System Data", "SourceInventory", 21, Empty
End Sub

Sub EnsureTable(ByVal wb, ByVal sheetName, ByVal tableName, ByVal headerRow, ByVal expectedHeaders)
    Dim ws, table, candidate, lastCol, lastRow, index, headerCount, sourceRange
    Set ws = wb.Worksheets(sheetName)
    Set table = Nothing

    On Error Resume Next
    Set table = ws.ListObjects(tableName)
    On Error GoTo 0

    If table Is Nothing Then
        For Each candidate In ws.ListObjects
            If candidate.HeaderRowRange.Row = headerRow Then
                Set table = candidate
                Exit For
            End If
        Next
    End If

    If IsArray(expectedHeaders) Then
        headerCount = UBound(expectedHeaders) - LBound(expectedHeaders) + 1
        ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow + 2, headerCount)).UnMerge
        For index = LBound(expectedHeaders) To UBound(expectedHeaders)
            ws.Cells(headerRow, index - LBound(expectedHeaders) + 1).Value = expectedHeaders(index)
        Next
        lastCol = headerCount
    Else
        lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    End If

    If lastCol < 1 Then Err.Raise vbObjectError + 110, "NMDC Setup", "No table headers found on " & sheetName

    If table Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow < headerRow + 1 Then lastRow = headerRow + 1
        Set sourceRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
        Set table = ws.ListObjects.Add(xlSrcRange, sourceRange, , xlYes)
        table.Name = tableName
        table.TableStyle = "TableStyleMedium2"
    ElseIf StrComp(table.Name, tableName, 1) <> 0 Then
        table.Name = tableName
    End If

    If table.ListRows.Count = 0 Then table.ListRows.Add
End Sub

Sub SetWorkbookConfig(ByVal wb, ByVal keyName, ByVal configValue)
    Dim ws, table, row, settingColumn, valueColumn, found
    Set ws = wb.Worksheets("Configuration")
    Set table = ws.ListObjects("Configuration")
    settingColumn = table.ListColumns("Setting").Index
    valueColumn = table.ListColumns("Current value").Index
    found = False

    For Each row In table.ListRows
        If StrComp(Trim(CStr(row.Range.Cells(1, settingColumn).Value)), keyName, 1) = 0 Then
            row.Range.Cells(1, valueColumn).Value = configValue
            found = True
            Exit For
        End If
    Next

    If Not found Then
        Set row = table.ListRows.Add
        row.Range.Cells(1, settingColumn).Value = keyName
        row.Range.Cells(1, valueColumn).Value = configValue
    End If
End Sub

Sub StyleHomeDashboard(ByVal wb)
    Dim ws
    Set ws = wb.Worksheets("Home")

    ws.Cells.Font.Name = "Aptos"
    ws.Range("A1:L33").Interior.Color = RGB(247, 249, 252)

    With ws.Range("A1:L1")
        .Interior.Color = RGB(11, 31, 51)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 22
        .RowHeight = 34
    End With
    With ws.Range("A2:L2")
        .Interior.Color = RGB(232, 241, 247)
        .Font.Color = RGB(65, 78, 92)
        .Font.Size = 11
        .RowHeight = 28
    End With
    With ws.Range("A4:L4")
        .Interior.Color = RGB(255, 247, 219)
        .Font.Color = RGB(122, 90, 0)
        .Font.Bold = True
        .RowHeight = 26
    End With

    StyleCard ws.Range("A5:C9")
    StyleCard ws.Range("D5:F9")
    StyleCard ws.Range("G5:I9")
    StyleCard ws.Range("J5:L9")

    On Error Resume Next
    ws.Range("A11:L11").UnMerge
    ws.Range("A11:L11").Merge
    On Error GoTo 0
    With ws.Range("A11:L11")
        .Value = "INDEX ACTIONS"
        .Interior.Color = RGB(20, 108, 148)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = -4108
        .RowHeight = 22
    End With

    On Error Resume Next
    ws.Range("A26:L26").UnMerge
    ws.Range("A26:L26").Merge
    On Error GoTo 0
    With ws.Range("A26:L26")
        .Value = "RECOVERY, RESET & HELP"
        .Interior.Color = RGB(91, 100, 112)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = -4108
        .RowHeight = 22
    End With

    With ws.Range("A29:L33")
        .Interior.Color = RGB(234, 244, 251)
        .Font.Color = RGB(31, 41, 55)
        .WrapText = True
    End With

    ws.Columns("A:L").ColumnWidth = 14
    ws.Rows("12:28").RowHeight = 24
    ws.Activate
    excel.ActiveWindow.DisplayGridlines = False
End Sub

Sub StyleCard(ByVal area)
    With area
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = 1
        .Borders.Color = RGB(216, 225, 232)
    End With
End Sub

Sub AttachHomeButtons(ByVal wb)
    Dim ws, ranges, macros, labels, fills, index, area, button, shape
    Set ws = wb.Worksheets("Home")

    For Each shape In ws.Shapes
        If Left(CStr(shape.Name), 12) = "NMDC_Action_" Then shape.Delete
    Next

    ranges = Array("A12:D13", "E12:H13", "I12:L13", _
                   "A15:D16", "E15:H16", "I15:L16", _
                   "A18:D19", "E18:H19", "I18:L19", _
                   "A21:D22", "E21:H22", "I21:L22", _
                   "A24:D25", "E24:H25", "I24:L25", _
                   "A27:D28", "E27:H28", "I27:L28")
    labels = Array("Update Changed Files", "Full Rescan / Rebuild All", "Review Pending Update", _
                   "Approve Update", "Hold Update", "Reject Update", _
                   "Select Data Folder", "Review Flags", "Save Review Decisions", _
                   "Configuration", "Rules & Mappings", "View Log", _
                   "Flag Wrong Data", "Report Requirement / Problem", "Refresh Dashboard", _
                   "Undo Last Approval", "Reset All Records", "Help")
    macros = Array("NMDC_UpdateChangedFiles", "NMDC_FullRescan", "NMDC_ReviewPendingUpdate", _
                   "NMDC_ApproveUpdate", "NMDC_HoldUpdate", "NMDC_RejectUpdate", _
                   "NMDC_SelectDataFolder", "NMDC_ReviewFlags", "NMDC_SaveReviewDecisions", _
                   "NMDC_OpenConfiguration", "NMDC_OpenRulesMappings", "NMDC_ViewLog", _
                   "NMDC_FlagWrongDataFromTable", "NMDC_ReportRequirementFromTable", "NMDC_RefreshDashboard", _
                   "NMDC_UndoLastApproval", "NMDC_ResetAllRecords", "NMDC_OpenHelp")
    fills = Array(RGB(20,108,148), RGB(20,108,148), RGB(194,139,0), _
                  RGB(46,125,50), RGB(194,139,0), RGB(198,40,40), _
                  RGB(20,108,148), RGB(194,139,0), RGB(46,125,50), _
                  RGB(20,108,148), RGB(20,108,148), RGB(91,100,112), _
                  RGB(20,108,148), RGB(91,100,112), RGB(46,125,50), _
                  RGB(194,139,0), RGB(198,40,40), RGB(91,100,112))

    For index = 0 To UBound(ranges)
        Set area = ws.Range(ranges(index))
        area.Hyperlinks.Delete
        area.Cells(1, 1).Value = labels(index)
        Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left + 2, area.Top + 2, area.Width - 4, area.Height - 4)
        button.Name = "NMDC_Action_" & CStr(index + 1)
        button.OnAction = macros(index)
        button.TextFrame.Characters.Text = labels(index)
        button.TextFrame.HorizontalAlignment = -4108
        button.TextFrame.VerticalAlignment = 3
        button.Fill.ForeColor.RGB = fills(index)
        button.Line.ForeColor.RGB = fills(index)
        button.TextFrame.Characters.Font.Name = "Aptos"
        button.TextFrame.Characters.Font.Size = 10
        button.TextFrame.Characters.Font.Bold = True
        button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        button.Placement = 1
    Next
End Sub

Sub ConfigureReviewFlags(ByVal wb)
    Dim ws, table, area, button, shape
    Set ws = wb.Worksheets("Review Flags")
    Set table = ws.ListObjects("ReviewFlags")

    On Error Resume Next
    ws.Range("A4:O4").UnMerge
    ws.Range("A4:O4").Merge
    On Error GoTo 0
    With ws.Range("A4:O4")
        .Value = "HOW TO REVIEW: 1) Open the Source File hyperlink. 2) Choose User Decision from the dropdown. 3) Add a User Comment only when explanation is needed. 4) Choose Resolution Status. 5) Click Save Review Decisions."
        .Interior.Color = RGB(255, 247, 219)
        .Font.Color = RGB(122, 90, 0)
        .Font.Bold = True
        .WrapText = True
        .RowHeight = 42
    End With

    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C:D").ColumnWidth = 48
    ws.Columns("H").ColumnWidth = 52
    ws.Columns("I").ColumnWidth = 28
    ws.Columns("L").ColumnWidth = 30
    ws.Columns("M").ColumnWidth = 42
    ws.Columns("N").ColumnWidth = 22

    For Each shape In ws.Shapes
        If CStr(shape.Name) = "NMDC_Save_Review" Then shape.Delete
    Next
    ws.Columns("P:S").ColumnWidth = 14
    Set area = ws.Range("P4:S5")
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Save_Review"
    button.OnAction = "NMDC_SaveReviewDecisions"
    button.TextFrame.Characters.Text = "Save Review Decisions"
    button.TextFrame.HorizontalAlignment = -4108
    button.TextFrame.VerticalAlignment = 3
    button.Fill.ForeColor.RGB = RGB(46, 125, 50)
    button.Line.ForeColor.RGB = RGB(46, 125, 50)
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 10
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255,255,255)
End Sub

Sub ApplyUserDropdowns(ByVal wb)
    ApplyTableDropdown wb.Worksheets("Review Flags").ListObjects("ReviewFlags"), "User Decision", _
        "ACKNOWLEDGED,NO ACTION REQUIRED,NEEDS SOURCE CORRECTION,NEEDS PARSER/MAPPING FIX,HOLD FOR REVIEW"
    ApplyTableDropdown wb.Worksheets("Review Flags").ListObjects("ReviewFlags"), "Resolution Status", _
        "OPEN,ACKNOWLEDGED,RESOLVED,DEFERRED"
    ApplyTableDropdown wb.Worksheets("User Decisions").ListObjects("UserDecisionLog"), "Decision Type", _
        "APPROVE,HOLD,REJECT,OVERRIDE,CONFIGURATION CHANGE"
    ApplyTableDropdown wb.Worksheets("Rules & Mappings").ListObjects("ClassificationRules"), "Enabled", "YES,NO"
    ApplyTableDropdown wb.Worksheets("Rules & Mappings").ListObjects("ClassificationRules"), "Include", "YES,NO"
End Sub

Sub ApplyTableDropdown(ByVal table, ByVal columnName, ByVal listValues)
    Dim target
    On Error Resume Next
    Set target = table.ListColumns(columnName).DataBodyRange
    On Error GoTo 0
    If target Is Nothing Then Exit Sub
    target.Validation.Delete
    target.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, listValues
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
End Sub

Sub NormalizeMergedUiRanges(ByVal wb)
    Dim ws, scanRange, cell, mergeArea, seen, address, topValue

    ' The source workbook contains many merged UI blocks. Some builder-generated
    ' files retain duplicate values in the hidden cells underneath a merge. Excel
    ' then warns on the next open that merging will discard those values. Normalize
    ' only the UI area (rows 1-40), never the extracted data tables below it.
    For Each ws In wb.Worksheets
        Set seen = CreateObject("Scripting.Dictionary")
        Set scanRange = ws.Range("A1:Z40")
        For Each cell In scanRange.Cells
            If cell.MergeCells Then
                Set mergeArea = cell.MergeArea
                address = mergeArea.Address
                If Not seen.Exists(address) Then
                    seen.Add address, True
                    topValue = mergeArea.Cells(1, 1).Value
                    mergeArea.UnMerge
                    ws.Range(address).ClearContents
                    ws.Range(address).Cells(1, 1).Value = topValue
                    ws.Range(address).Merge
                End If
            End If
        Next
    Next
End Sub

Sub ShowFailure(ByVal friendlyMessage)
    Dim detail
    detail = Err.Number & " - " & Err.Description
    On Error Resume Next
    If IsObject(workbook) Then workbook.Close False
    If IsObject(excel) Then excel.Quit
    MsgBox friendlyMessage & vbCrLf & vbCrLf & detail, vbCritical, "NMDC Document Index Setup"
End Sub

Function SafeTimestamp()
    Dim value
    value = Now
    SafeTimestamp = Year(value) & Right("0" & Month(value), 2) & Right("0" & Day(value), 2) & _
                    "-" & Right("0" & Hour(value), 2) & Right("0" & Minute(value), 2) & Right("0" & Second(value), 2)
End Function

Function RGB(ByVal redValue, ByVal greenValue, ByVal blueValue)
    RGB = CLng(redValue) + (CLng(greenValue) * 256) + (CLng(blueValue) * 65536)
End Function
