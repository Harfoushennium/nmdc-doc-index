$excelSafetyType = 'BathyExcelWindowProcess'
if (-not ($excelSafetyType -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class BathyExcelWindowProcess
{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@ | Out-Null
}

function Get-ExcelAutomationIdentity {
    param([Parameter(Mandatory=$true)][object]$Application)
    $window = [IntPtr]([long]$Application.Hwnd)
    if ($window -eq [IntPtr]::Zero) { throw 'Excel did not expose a valid application window handle.' }
    [uint32]$processId = 0
    $threadId = [BathyExcelWindowProcess]::GetWindowThreadProcessId($window, [ref]$processId)
    if ($threadId -eq 0 -or $processId -eq 0) { throw 'Could not resolve Excel HWND to a process identity.' }
    $process = Get-Process -Id ([int]$processId) -ErrorAction Stop
    [pscustomobject]@{ ProcessId = [int]$processId; StartTimeUtcTicks = $process.StartTime.ToUniversalTime().Ticks }
}

function Get-ExcelProcessInventory {
    $inventory = @{}
    foreach ($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        $inventory[[int]$process.Id] = $process.StartTime.ToUniversalTime().Ticks
    }
    return $inventory
}

function New-OwnedExcelApplication {
    param([object]$Application = $null)
    $factoryStartedUtcTicks = [DateTime]::UtcNow.Ticks
    $existing = Get-ExcelProcessInventory
    $createdHere = $null -eq $Application
    if ($createdHere) { $Application = New-Object -ComObject Excel.Application }
    try {
        $identity = Get-ExcelAutomationIdentity -Application $Application
        if ($identity.StartTimeUtcTicks -lt $factoryStartedUtcTicks) {
            throw "Excel PID $($identity.ProcessId) started before this automation factory; refusing to claim it as owned."
        }
        if ($existing.ContainsKey($identity.ProcessId) -and $existing[$identity.ProcessId] -eq $identity.StartTimeUtcTicks) {
            throw "Excel automation resolved to preexisting PID $($identity.ProcessId); no workbook or application settings were changed."
        }
        $owner = [pscustomobject]@{
            Application = $Application
            ProcessId = $identity.ProcessId
            StartTimeUtcTicks = $identity.StartTimeUtcTicks
            PreexistingProcesses = $existing
            CreatedHere = $createdHere
            Workbooks = New-Object System.Collections.ArrayList
        }
        Assert-OwnedExcelApplication -Owner $owner
        return $owner
    } catch {
        if ($createdHere -and $Application) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Application) }
        throw
    }
}

function Assert-OwnedExcelApplication {
    param([Parameter(Mandatory=$true)][object]$Owner)
    if (-not $Owner.CreatedHere) { throw 'The Excel application is not owned by this automation run.' }
    $identity = Get-ExcelAutomationIdentity -Application $Owner.Application
    if ($identity.ProcessId -ne $Owner.ProcessId -or $identity.StartTimeUtcTicks -ne $Owner.StartTimeUtcTicks) {
        throw 'Excel HWND no longer resolves to the owned process identity; refusing the operation.'
    }
    if ($Owner.PreexistingProcesses.ContainsKey($identity.ProcessId) -and $Owner.PreexistingProcesses[$identity.ProcessId] -eq $identity.StartTimeUtcTicks) {
        throw 'Excel resolved to a protected preexisting process; refusing the operation.'
    }
}

function New-OwnedExcelWorkbook {
    param([Parameter(Mandatory=$true)][object]$Owner)
    Assert-OwnedExcelApplication -Owner $Owner
    $workbook = $Owner.Application.Workbooks.Add()
    Assert-OwnedExcelApplication -Owner $Owner
    [void]$Owner.Workbooks.Add($workbook)
    return $workbook
}

function Open-OwnedExcelWorkbook {
    param(
        [Parameter(Mandatory=$true)][object]$Owner,
        [Parameter(Mandatory=$true)][string]$Path,
        [int]$UpdateLinks = 0,
        [bool]$ReadOnly = $false
    )
    Assert-OwnedExcelApplication -Owner $Owner
    $workbook = $Owner.Application.Workbooks.Open($Path, $UpdateLinks, $ReadOnly)
    Assert-OwnedExcelApplication -Owner $Owner
    [void]$Owner.Workbooks.Add($workbook)
    return $workbook
}

function Save-OwnedExcelWorkbook {
    param([Parameter(Mandatory=$true)][object]$Owner, [Parameter(Mandatory=$true)][object]$Workbook)
    Assert-OwnedExcelWorkbook -Owner $Owner -Workbook $Workbook
    $Workbook.Save()
    Assert-OwnedExcelWorkbook -Owner $Owner -Workbook $Workbook
}

function SaveAs-OwnedExcelWorkbook {
    param([Parameter(Mandatory=$true)][object]$Owner, [Parameter(Mandatory=$true)][object]$Workbook, [Parameter(Mandatory=$true)][string]$Path, [int]$FileFormat = 52)
    Assert-OwnedExcelWorkbook -Owner $Owner -Workbook $Workbook
    $Workbook.SaveAs($Path, $FileFormat)
    Assert-OwnedExcelWorkbook -Owner $Owner -Workbook $Workbook
}

function Assert-OwnedExcelWorkbook {
    param([Parameter(Mandatory=$true)][object]$Owner, [Parameter(Mandatory=$true)][object]$Workbook)
    Assert-OwnedExcelApplication -Owner $Owner
    if (-not $Owner.Workbooks.Contains($Workbook)) { throw 'Workbook reference is not tracked by this automation owner.' }
    $identity = Get-ExcelAutomationIdentity -Application $Workbook.Application
    if ($identity.ProcessId -ne $Owner.ProcessId -or $identity.StartTimeUtcTicks -ne $Owner.StartTimeUtcTicks) {
        throw 'Workbook belongs to a different Excel process identity; refusing the operation.'
    }
}

function Close-OwnedExcelWorkbook {
    param([Parameter(Mandatory=$true)][object]$Owner, [Parameter(Mandatory=$true)][object]$Workbook, [bool]$SaveChanges = $false)
    Assert-OwnedExcelWorkbook -Owner $Owner -Workbook $Workbook
    $Workbook.Close($SaveChanges)
    Assert-OwnedExcelApplication -Owner $Owner
    [void]$Owner.Workbooks.Remove($Workbook)
}

function Close-OwnedExcelApplication {
    param([Parameter(Mandatory=$true)][object]$Owner)
    if (-not $Owner.CreatedHere) { return }
    try {
        foreach ($workbook in @($Owner.Workbooks)) {
            try {
                Close-OwnedExcelWorkbook -Owner $Owner -Workbook $workbook -SaveChanges $false
                if ([Runtime.InteropServices.Marshal]::IsComObject($workbook)) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
            } catch { Write-Warning $_.Exception.Message }
        }
        Assert-OwnedExcelApplication -Owner $Owner
        if ($Owner.Application.Workbooks.Count -eq 0) {
            Assert-OwnedExcelApplication -Owner $Owner
            $Owner.Application.Quit()
        } else {
            Write-Warning 'Excel has untracked workbooks open; leaving the owned application running rather than closing them.'
        }
    } catch {
        Write-Warning "Excel cleanup stopped because ownership could not be verified: $($_.Exception.Message)"
    } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Owner.Application)
    }
}
