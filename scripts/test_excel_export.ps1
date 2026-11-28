$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel_automation_safety.ps1')
$repository=(Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$fixture=(Resolve-Path (Join-Path $repository 'sandbox/excel-proof')).Path
if (-not $fixture.StartsWith($repository+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture outside workspace' }
$owner=$null
try {
    $owner=New-OwnedExcelApplication
    Assert-OwnedExcelApplication -Owner $owner
    $owner.Application.Visible=$false
    $document=Open-OwnedExcelWorkbook -Owner $owner -Path (Join-Path $fixture 'document.xlsx') -ReadOnly $true
    $source=Open-OwnedExcelWorkbook -Owner $owner -Path (Join-Path $fixture 'source.xlsx') -ReadOnly $true
    $index=Open-OwnedExcelWorkbook -Owner $owner -Path (Join-Path $fixture 'index.xlsx') -ReadOnly $true
    Assert-OwnedExcelWorkbook -Owner $owner -Workbook $index
    if ($index.Worksheets.Count -ne 1 -or $index.Worksheets.Item(1).ListObjects.Count -ne 1) { throw 'INDEX/table count mismatch' }
    $sheet=$index.Worksheets.Item(1)
    if ($sheet.Cells.Item(2,2).Value2 -cne '00001') { throw 'Leading zeros lost' }
    if ($sheet.Hyperlinks.Count -ne 2) { throw 'Native hyperlink count mismatch' }
    $sheet.Hyperlinks.Item(1).Follow()
    if ($owner.Application.ActiveWorkbook.Name -ne $document.Name) { throw 'Document hyperlink did not activate the target' }
    $index.Activate()
    $sheet.Hyperlinks.Item(2).Follow()
    if ($owner.Application.ActiveWorkbook.Name -ne $source.Name) { throw 'Source hyperlink did not activate the target' }
    $empty=Open-OwnedExcelWorkbook -Owner $owner -Path (Join-Path $fixture 'empty.xlsx') -ReadOnly $true
    if ($empty.Worksheets.Item(1).ListObjects.Count -ne 1 -or $empty.Worksheets.Item(1).UsedRange.Rows.Count -gt 2) { throw 'Empty INDEX table invalid' }
    $dates=Open-OwnedExcelWorkbook -Owner $owner -Path (Join-Path $fixture 'dates.xlsx') -ReadOnly $true
    $dateSheet=$dates.Worksheets.Item(1)
    if ($dateSheet.Cells.Item(2,1).Value2 -cne '1899-12-31') { throw 'Pre-1900 literal date changed' }
    foreach ($check in @(@(3,'01/01/1900'),@(4,'28/02/1900'),@(5,'01/03/1900'))) {
        if ($dateSheet.Cells.Item($check[0],1).Text -cne $check[1]) { throw 'Native Excel date boundary changed' }
    }
    @{ status='PASS'; excel_version=$owner.Application.Version; excel_build=$owner.Application.Build; fixture='synthetic local XLSX'; normal_open=$true; document_follow=$true; source_follow=$true; pre1900_literal=$true; native_1900_boundaries=$true; physical_owner_uat='NOT RUN'; machine='development PC' } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $fixture 'evidence.json') -Encoding utf8
} finally {
    if ($owner) { Close-OwnedExcelApplication -Owner $owner }
}
