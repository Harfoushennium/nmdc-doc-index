Option Explicit

Const xlOpenXMLWorkbookMacroEnabled = 52
Const xlSrcRange = 1
Const xlYes = 1
Const xlUp = -4162
Const xlToLeft = -4159
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
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Refresh.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Actions.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_TableActions.bas"), "A required Excel table-action module is missing."
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
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Refresh.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Actions.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_TableActions.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Rules.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Startup.bas")
If Err.Number <> 0 Then
    ShowFailure "Excel could not attach the production actions."
    WScript.Quit 8
End If

AttachHomeButtons workbook
SetWorkbookConfig workbook, "Engine Executable Path", enginePath
SetWorkbookConfig workbook, "Runtime Folder", runtimeFolder
SetWorkbookConfig workbook, "Configuration Folder", configFolder
SetWorkbookConfig workbook, "Classification Rules File", fso.BuildPath(configFolder, "classification_rules.csv")
SetWorkbookConfig workbook, "Project Identity Overrides File", fso.BuildPath(configFolder, "project_identity_overrides.csv")
workbook.Worksheets("System Data").Visible = 2
workbook.Save
workbook.Close True
excel.Quit
On Error GoTo 0

MsgBox "NMDC_Document_Index.xlsm was created successfully." & vbCrLf & vbCrLf & _
       "Open it in Microsoft Excel, enable macros, select the DATA folder, and press Update Changed Files.", _
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

Sub AttachHomeButtons(ByVal wb)
    Dim ws, ranges, macros, index, area, button, label
    Set ws = wb.Worksheets("Home")
    ranges = Array("A12:D13", "E12:H13", "I12:L13", _
                   "A15:D16", "E15:H16", "I15:L16", _
                   "A18:D19", "E18:H19", "I18:L19", _
                   "A21:D22", "E21:H22", "I21:L22", _
                   "A24:D25", "E24:H25", "I24:L25")
    macros = Array("NMDC_UpdateChangedFiles", "NMDC_FullRescan", "NMDC_ReviewPendingUpdate", _
                   "NMDC_ApproveUpdate", "NMDC_HoldUpdate", "NMDC_RejectUpdate", _
                   "NMDC_SelectDataFolder", "NMDC_ReviewFlags", "NMDC_FlagWrongDataFromTable", _
                   "NMDC_OpenConfiguration", "NMDC_ReportRequirementFromTable", "NMDC_ViewLog", _
                   "NMDC_RefreshDashboard", "NMDC_OpenRulesMappings", "NMDC_OpenHelp")

    For index = 0 To UBound(ranges)
        Set area = ws.Range(ranges(index))
        label = CStr(area.Cells(1, 1).Value)
        area.Hyperlinks.Delete
        Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width, area.Height)
        button.Name = "NMDC_Action_" & CStr(index + 1)
        button.OnAction = macros(index)
        button.TextFrame.Characters.Text = label
        button.TextFrame.HorizontalAlignment = -4108
        button.TextFrame.VerticalAlignment = 3
        button.Fill.ForeColor.RGB = area.Cells(1, 1).Interior.Color
        button.Line.ForeColor.RGB = RGB(180, 195, 205)
        button.TextFrame.Characters.Font.Name = "Aptos"
        button.TextFrame.Characters.Font.Size = 10
        button.TextFrame.Characters.Font.Bold = True
        button.TextFrame.Characters.Font.Color = RGB(20, 108, 148)
        button.Placement = 1
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
