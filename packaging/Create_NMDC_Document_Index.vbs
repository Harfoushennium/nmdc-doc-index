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
Dim modulesFolder, enginePath, runtimeFolder, configFolder, response, localAppData, appDataRoot
Dim setupStage, setupObject
setupStage = "startup"
setupObject = ""

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
packageRoot = fso.GetParentFolderName(WScript.ScriptFullName)
sourceWorkbook = fso.BuildPath(packageRoot, "source\NMDC_Document_Index_Base.xlsx")
outputWorkbook = fso.BuildPath(packageRoot, "NMDC_Document_Index.xlsm")
modulesFolder = fso.BuildPath(packageRoot, "vba")
enginePath = fso.BuildPath(packageRoot, "engine\nmdc_index_engine.exe")
configFolder = fso.BuildPath(packageRoot, "config")

' Keep runtime/state/cache out of the extracted package so normal scans do not
' create thousands of sync events when the package itself sits under OneDrive.
localAppData = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
If Len(localAppData) = 0 Or InStr(localAppData, "%LOCALAPPDATA%") > 0 Then localAppData = packageRoot
appDataRoot = fso.BuildPath(localAppData, "NMDC Document Index")
runtimeFolder = fso.BuildPath(appDataRoot, "runtime")
EnsureFolderTree appDataRoot
EnsureFolderTree runtimeFolder

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
RequireFile fso.BuildPath(configFolder, "source_exclusions.csv"), "The source-selection file is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Engine.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Csv.bas"), "A required safe CSV module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Refresh.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Actions.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_TableActions.bas"), "A required Excel table-action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Admin.bas"), "A required Excel administration module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Rules.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Startup.bas"), "A required Excel action module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Performance.bas"), "The responsive scan module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_OwnerUX.bas"), "The owner guidance module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_Checkboxes.bas"), "The source-selection checkbox module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_LiveFilter.bas"), "The Live Filter module is missing."
RequireFile fso.BuildPath(modulesFolder, "Cls_LiveFilter_Listener.cls"), "The Live Filter listener class is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_CustomFields.bas"), "The Custom Fields module is missing."
RequireFile fso.BuildPath(modulesFolder, "modNMDC_CustomFieldsSetup.bas"), "The Custom Fields setup module is missing."

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
TraceStep "workbook-saved-as-xlsm"

Err.Clear
' SaveAs can leave a stale COM proxy in a long-running Excel instance.
' Reacquire the exact saved workbook: excel.Workbooks(fso.GetFileName(outputWorkbook))
Set workbook = ReacquireSavedWorkbook(excel, outputWorkbook)
If workbook Is Nothing Then
    Err.Raise vbObjectError + 111, "NMDC Setup", "Saved workbook could not be reacquired after SaveAs: " & outputWorkbook
End If
If Err.Number <> 0 Then
    ShowFailure "Excel created the XLSM but could not reacquire the saved workbook safely."
    WScript.Quit 6
End If
TraceStep "workbook-reacquired-after-saveas"

Err.Clear
setupStage = "EnsureNamedTables"
setupObject = "initializing"
EnsureNamedTables workbook
If Err.Number <> 0 Then
    Dim namedTableErrorNumber, namedTableErrorDescription
    namedTableErrorNumber = Err.Number
    namedTableErrorDescription = Err.Description
    TraceStep "setup-error stage=" & setupStage & " object=" & setupObject & " error=" & CStr(namedTableErrorNumber) & " - " & namedTableErrorDescription
    ShowFailure "Excel could not validate the required named tables." & vbCrLf & _
                "Stage: " & setupStage & vbCrLf & _
                "Object: " & setupObject & vbCrLf & _
                "Error: " & CStr(namedTableErrorNumber) & " - " & namedTableErrorDescription
    WScript.Quit 7
End If
TraceStep "named-tables-ready"

ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Engine.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Csv.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Refresh.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Actions.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_TableActions.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Admin.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Rules.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Startup.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Performance.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_OwnerUX.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_Checkboxes.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_LiveFilter.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "Cls_LiveFilter_Listener.cls")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_CustomFields.bas")
ImportModule workbook, fso.BuildPath(modulesFolder, "modNMDC_CustomFieldsSetup.bas")
If Err.Number <> 0 Then
    ShowFailure "Excel could not attach the production actions."
    WScript.Quit 8
End If
TraceStep "vba-modules-imported"

ConfigureFastStartup workbook
If Err.Number <> 0 Then
    ShowFailure "Excel could not configure the fast workbook startup."
    WScript.Quit 9
End If
TraceStep "fast-startup-configured"

Err.Clear
ConfigureOwnerEvents workbook
If Err.Number <> 0 Then
    ShowFailure "Excel could not configure the workbook interaction events."
    WScript.Quit 9
End If
TraceStep "owner-events-configured"

SetWorkbookConfig workbook, "Engine Executable Path", enginePath
SetWorkbookConfig workbook, "Runtime Folder", runtimeFolder
SetWorkbookConfig workbook, "Configuration Folder", configFolder
SetWorkbookConfig workbook, "Classification Rules File", fso.BuildPath(configFolder, "classification_rules.csv")
SetWorkbookConfig workbook, "Project Identity Overrides File", fso.BuildPath(configFolder, "project_identity_overrides.csv")
SetWorkbookConfig workbook, "Source Exclusions File", fso.BuildPath(configFolder, "source_exclusions.csv")
TraceStep "configuration-paths-written"

StyleHomeDashboard workbook
AttachHomeButtons workbook
ConfigureReviewFlags workbook
ApplyUserDropdowns workbook
workbook.Worksheets("System Data").Visible = 2
NormalizeMergedUiRanges workbook
TraceStep "base-ui-normalized"

Err.Clear
excel.Run "'" & workbook.Name & "'!NMDC_EnsureCustomFieldsStructure"
If Err.Number <> 0 Then
    ShowFailure "Excel could not create the Custom Fields & Keywords workspace."
    WScript.Quit 10
End If
TraceStep "custom-fields-structure-ready"

Err.Clear
excel.Run "'" & workbook.Name & "'!NMDC_CustomFieldsInitialize"
If Err.Number <> 0 Then
    ShowFailure "Excel could not initialize Custom Fields & Keywords."
    WScript.Quit 10
End If
TraceStep "custom-fields-initialized"

Err.Clear
excel.Run "'" & workbook.Name & "'!NMDC_LiveFilterInitialize"
If Err.Number <> 0 Then
    ShowFailure "Excel could not initialize the Live Filter."
    WScript.Quit 10
End If
TraceStep "live-filter-initialized"

Err.Clear
excel.Run "'" & workbook.Name & "'!NMDC_ApplyOwnerUX"
If Err.Number <> 0 Then
    ShowFailure "Excel could not apply the final owner guidance."
    WScript.Quit 10
End If
TraceStep "owner-ux-applied"
workbook.Save
TraceStep "final-workbook-saved"
workbook.Close True
excel.Quit
On Error GoTo 0

MsgBox "NMDC_Document_Index.xlsm was created successfully." & vbCrLf & vbCrLf & _
       "Runtime/cache: " & runtimeFolder & vbCrLf & vbCrLf & _
       "Open the workbook, enable macros, select the DATA folder, and use Update Changed Files for normal work. Full Rescan is intended only for deliberate rebuilds.", _
       vbInformation, "NMDC Document Index Setup"

Function ReacquireSavedWorkbook(ByVal excelApp, ByVal expectedPath)
    Dim candidate, expectedName, candidatePath
    Set ReacquireSavedWorkbook = Nothing
    expectedName = fso.GetFileName(expectedPath)

    On Error Resume Next
    Err.Clear
    Set candidate = excelApp.Workbooks(expectedName)
    If Not candidate Is Nothing Then
        Set ReacquireSavedWorkbook = candidate
        Err.Clear
        Exit Function
    End If
    Err.Clear

    For Each candidate In excelApp.Workbooks
        candidatePath = ""
        candidatePath = candidate.FullName
        If Err.Number = 0 Then
            If StrComp(fso.GetAbsolutePathName(candidatePath), fso.GetAbsolutePathName(expectedPath), 1) = 0 Then
                Set ReacquireSavedWorkbook = candidate
                Exit Function
            End If
        End If
        Err.Clear
    Next
    On Error GoTo 0
End Function

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

Sub EnsureFolderTree(ByVal folderPath)
    Dim parentPath
    If Len(folderPath) = 0 Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub
    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then EnsureFolderTree parentPath
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Sub EnsureNamedTables(ByVal wb)
    EnsureSourceSelectionSheet wb
    EnsureTable wb, "Master Documents", "MasterDocuments", 5, Empty
    EnsureTable wb, "Revisions", "RevisionRegister", 5, Empty
    EnsureTable wb, "Transactions", "EventRegister", 5, Empty
    EnsureTable wb, "Pending Update", "PendingUpdate", 5, Array("Change Type", "Project No.", "Document No.", "Revision", "Event Type", "Plain-English Summary", "Source File", "Review Required", "Record Identity")
    EnsureTable wb, "Review Flags", "ReviewFlags", 5, Empty
    EnsureTable wb, "Source Selection", "SourceSelection", 5, Array("Include in Index?", "Source File", "Source Family", "Current Status", "Owner Note", "Selection Reason", "Last Processed Run")
    EnsureTable wb, "User Decisions", "UserDecisionLog", 11, Empty
    EnsureTable wb, "Configuration", "Configuration", 5, Empty
    EnsureTable wb, "Rules & Mappings", "ClassificationRules", 5, Empty
    EnsureTable wb, "Update History", "UpdateHistory", 5, Empty
    EnsureTable wb, "Error Log", "ErrorLog", 5, Array("Date/Time", "Severity", "Action", "Plain-English Error", "Recommended Action", "Technical Detail", "Run ID", "Source File", "Worksheet", "Source Row/Cell")
    EnsureTable wb, "System Data", "BaselineCounts", 5, Empty
    EnsureTable wb, "System Data", "SourceInventory", 21, Empty
End Sub

Sub EnsureSourceSelectionSheet(ByVal wb)
    Dim ws
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Source Selection")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Source Selection"
        ws.Range("A1").Value = "Source Selection"
        ws.Range("A2").Value = "Use the checkboxes to choose which source workbooks participate in the index. This does not modify or delete the source files."
    End If
End Sub

Sub EnsureTable(ByVal wb, ByVal sheetName, ByVal tableName, ByVal headerRow, ByVal expectedHeaders)
    Dim ws, table, candidate, lastCol, lastRow, index, headerCount, sourceRange
    TraceStep "ensure-table-start " & sheetName & "!" & tableName
    setupStage = "EnsureTable"
    setupObject = sheetName & "!" & tableName
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
    TraceStep "ensure-table-ok " & sheetName & "!" & tableName
End Sub

Sub ConfigureFastStartup(ByVal wb)
    Dim component, codeModule, lineNo, lineText, thisComponent, thisModule, eventCode, sourceText

    ' Keep the fast startup introduced after the last working owner build, but do
    ' not assume that ThisWorkbook already contains code. A fresh base workbook
    ' can have CountOfLines = 0; calling Lines(1, 0) raises -2147024809.
    Set component = wb.VBProject.VBComponents("modNMDC_Startup")
    Set codeModule = component.CodeModule
    For lineNo = 1 To codeModule.CountOfLines
        lineText = Trim(codeModule.Lines(lineNo, 1))
        If StrComp(lineText, "Public Sub Auto_Open()", 1) = 0 Then
            codeModule.ReplaceLine lineNo, "Public Sub NMDC_LegacyAutoOpen()"
            Exit For
        End If
    Next

    Set thisComponent = wb.VBProject.VBComponents(wb.CodeName)
    Set thisModule = thisComponent.CodeModule
    sourceText = ""
    If thisModule.CountOfLines > 0 Then
        sourceText = thisModule.Lines(1, thisModule.CountOfLines)
    End If

    If InStr(1, sourceText, "Private Sub Workbook_Open", 1) = 0 Then
        eventCode = vbCrLf & "Private Sub Workbook_Open()" & vbCrLf & _
                    "    On Error Resume Next" & vbCrLf & _
                    "    NMDC_FastStartup" & vbCrLf & _
                    "    On Error GoTo 0" & vbCrLf & _
                    "End Sub" & vbCrLf
        thisModule.AddFromString eventCode
    End If
End Sub

Sub ConfigureOwnerEvents(ByVal wb)
    Dim thisComponent, thisModule, sourceText, eventCode

    Set thisComponent = wb.VBProject.VBComponents(wb.CodeName)
    Set thisModule = thisComponent.CodeModule
    sourceText = ""
    If thisModule.CountOfLines > 0 Then
        sourceText = thisModule.Lines(1, thisModule.CountOfLines)
    End If

    eventCode = ""

    If InStr(1, sourceText, "Private Sub Workbook_SheetChange", 1) = 0 Then
        eventCode = eventCode & vbCrLf & _
            "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf & _
            "    On Error Resume Next" & vbCrLf & _
            "    NMDC_LiveFilterSheetChange Sh, Target" & vbCrLf & _
            "    On Error GoTo 0" & vbCrLf & _
            "End Sub" & vbCrLf
    End If

    If InStr(1, sourceText, "Private Sub Workbook_SheetActivate", 1) = 0 Then
        eventCode = eventCode & vbCrLf & _
            "Private Sub Workbook_SheetActivate(ByVal Sh As Object)" & vbCrLf & _
            "    On Error Resume Next" & vbCrLf & _
            "    NMDC_LiveFilterSheetActivate Sh" & vbCrLf & _
            "    NMDC_LiveFilterSelectionChange Sh, Target" & vbCrLf & _
            "    NMDC_EnsureCustomFieldsCurrent Sh" & vbCrLf & _
            "    On Error GoTo 0" & vbCrLf & _
            "End Sub" & vbCrLf
    End If

    If InStr(1, sourceText, "Private Sub Workbook_SheetSelectionChange", 1) = 0 Then
        eventCode = eventCode & vbCrLf & _
            "Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf & _
            "    On Error Resume Next" & vbCrLf & _
            "    NMDC_EnsureCustomFieldsCurrent Sh" & vbCrLf & _
            "    On Error GoTo 0" & vbCrLf & _
            "End Sub" & vbCrLf
    End If

    If Len(eventCode) > 0 Then thisModule.AddFromString eventCode
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

    ws.Range("A1:L42").Font.Name = "Aptos"
    ws.Range("A1:L42").Interior.Color = RGB(247, 249, 252)

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
    ws.Range("A11:L11").ClearContents
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
    ws.Range("A26:L26").ClearContents
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

    ws.Columns("A:L").ColumnWidth = 14
    ws.Rows("12:28").RowHeight = 24
    ws.Rows("40:41").RowHeight = 24
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
                   "A27:D28", "E27:H28", "I27:L28", _
                   "A40:D41")
    labels = Array("Update Changed Files", "Full Rescan / Rebuild All", "Review Pending Update", _
                   "Approve Update", "Hold Update", "Reject Update", _
                   "Select Data Folder", "Review Flags", "Save Review Decisions", _
                   "Configuration", "Rules & Mappings", "View Log", _
                   "Flag Wrong Data", "Report Requirement / Problem", "Refresh Dashboard", _
                   "Undo Last Approval", "Reset All Records", "Help", _
                   "Source Selection")
    macros = Array("NMDC_UpdateChangedFilesFast", "NMDC_FullRescanFast", "NMDC_ReviewPendingUpdateFast", _
                   "NMDC_ApproveUpdate", "NMDC_HoldUpdate", "NMDC_RejectUpdate", _
                   "NMDC_SelectDataFolder", "NMDC_ReviewFlagsFast", "NMDC_SaveReviewDecisions", _
                   "NMDC_OpenConfiguration", "NMDC_OpenRulesMappingsOwner", "NMDC_ViewLog", _
                   "NMDC_FlagWrongDataFromTable", "NMDC_ReportRequirementFromTable", "NMDC_RefreshDashboardFast", _
                   "NMDC_UndoLastApproval", "NMDC_ResetAllRecords", "NMDC_OpenHelp", _
                   "NMDC_OpenSourceSelection")
    fills = Array(RGB(20,108,148), RGB(194,139,0), RGB(194,139,0), _
                  RGB(46,125,50), RGB(194,139,0), RGB(198,40,40), _
                  RGB(20,108,148), RGB(194,139,0), RGB(46,125,50), _
                  RGB(20,108,148), RGB(20,108,148), RGB(91,100,112), _
                  RGB(20,108,148), RGB(91,100,112), RGB(46,125,50), _
                  RGB(194,139,0), RGB(198,40,40), RGB(91,100,112), _
                  RGB(20,108,148))

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
    ws.Range("A4:O4").ClearContents
    ws.Range("A4:O4").Merge
    On Error GoTo 0
    With ws.Range("A4:O4")
        .Value = "HOW TO REVIEW: 1) Open the Source File hyperlink. 2) For extraction/layout problems, select the row and click Request Parser / Mapping Fix; describe the expected layout. 3) The workbook creates a fix-request file but never rewrites its own parser silently. 4) After a corrected parser/config is installed, click Retry After Fix. Other decisions can still be saved with Save Review Decisions."
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
        If CStr(shape.Name) = "NMDC_Save_Review" Or _
           CStr(shape.Name) = "NMDC_Request_Review_Fix" Or _
           CStr(shape.Name) = "NMDC_Retry_Review_Fix" Then shape.Delete
    Next
    ws.Columns("P:S").ColumnWidth = 14

    Set area = ws.Range("P4:S4")
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Save_Review"
    button.OnAction = "NMDC_SaveReviewDecisions"
    button.TextFrame.Characters.Text = "Save Review Decisions"
    button.TextFrame.HorizontalAlignment = -4108
    button.TextFrame.VerticalAlignment = 3
    button.Fill.ForeColor.RGB = RGB(46, 125, 50)
    button.Line.ForeColor.RGB = RGB(46, 125, 50)
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 9
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255,255,255)

    Set area = ws.Range("P5:S5")
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Request_Review_Fix"
    button.OnAction = "NMDC_RequestParserMappingFix"
    button.TextFrame.Characters.Text = "Request Parser / Mapping Fix"
    button.TextFrame.HorizontalAlignment = -4108
    button.TextFrame.VerticalAlignment = 3
    button.Fill.ForeColor.RGB = RGB(194, 139, 0)
    button.Line.ForeColor.RGB = RGB(194, 139, 0)
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 9
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255,255,255)

    Set area = ws.Range("P6:S6")
    Set button = ws.Shapes.AddShape(msoShapeRoundedRectangle, area.Left, area.Top, area.Width, area.Height)
    button.Name = "NMDC_Retry_Review_Fix"
    button.OnAction = "NMDC_RetryAfterParserMappingFix"
    button.TextFrame.Characters.Text = "Retry After Fix"
    button.TextFrame.HorizontalAlignment = -4108
    button.TextFrame.VerticalAlignment = 3
    button.Fill.ForeColor.RGB = RGB(20, 108, 148)
    button.Line.ForeColor.RGB = RGB(20, 108, 148)
    button.TextFrame.Characters.Font.Name = "Aptos"
    button.TextFrame.Characters.Font.Size = 9
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

    For Each ws In wb.Worksheets
        Set seen = CreateObject("Scripting.Dictionary")
        Set scanRange = ws.Range("A1:Z45")
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
    Dim detail, logFile
    detail = Err.Number & " - " & Err.Description
    On Error Resume Next
    Set logFile = fso.CreateTextFile(fso.BuildPath(packageRoot, "setup_error.log"), True, True)
    logFile.WriteLine friendlyMessage
    logFile.WriteLine detail
    logFile.Close
    On Error Resume Next
    If IsObject(workbook) Then workbook.Close False
    If IsObject(excel) Then excel.Quit
    MsgBox friendlyMessage & vbCrLf & vbCrLf & detail, vbCritical, "NMDC Document Index Setup"
End Sub

Sub TraceStep(ByVal message)
    Dim traceFile
    On Error Resume Next
    Set traceFile = fso.OpenTextFile(fso.BuildPath(packageRoot, "setup_trace.log"), 8, True, -1)
    traceFile.WriteLine SafeTimestamp() & " " & message
    traceFile.Close
    On Error GoTo 0
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
