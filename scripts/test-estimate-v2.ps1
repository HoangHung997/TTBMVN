param(
    [string]$WorkbookPath = 'Y:\Mau DuToan\Du toan RPBM HoaLuNamDinh_Ver1.xlsx'
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'excel-test-process.ps1')
$runRoot = Join-Path $repoRoot ('tmp\V2-runtime-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$workingCopy = Join-Path $runRoot 'HoaLuNamDinh-V2-test.xlsx'
function Get-SharedHash([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','') }
    finally { $sha.Dispose(); $stream.Dispose() }
}
$sourceHash = Get-SharedHash $WorkbookPath
Copy-Item -LiteralPath $WorkbookPath -Destination $workingCopy
[void][Reflection.Assembly]::LoadFrom((Join-Path $repoRoot 'ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll'))
$addin = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot 'ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll'))
function Invoke-V2([string]$TypeName, [string]$MethodName, [object[]]$Arguments) {
    $type = $addin.GetType('ExcelAddIn1.Funtion.' + $TypeName, $true)
    $method = $type.GetMethod($MethodName)
    $values = New-Object 'object[]' $Arguments.Length
    for ($i=0; $i -lt $Arguments.Length; $i++) {
        if ($null -ne $Arguments[$i]) { $values[$i] = $Arguments[$i].PSObject.BaseObject }
    }
    return $method.Invoke($null, $values)
}
function Assert-V2([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.EnableEvents = $false
$excel.AutomationSecurity = 3
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $workbook.Save()
    Invoke-V2 'WorkbookProjectProfileService' 'Clear' @($workbook) | Out-Null
    $available = Invoke-V2 'RegulationPackageBootstrapService' 'LoadAvailablePackages' @()
    $target = @($available | Where-Object PackageId -eq 'BQP-RPBM-2025')[0]
    Assert-V2 ($null -ne $target) 'Package BQP-RPBM-2025 unavailable'
    $preview = Invoke-V2 'WorkbookEstimateV2PackageService' 'Preview' @($workbook,$target,(Get-Date),(Get-Date))
    Assert-V2 $preview.Plan.CanApply 'Initial package preview blocked'
    $raw = New-Object 'object[]' 2
    $raw[0]=$workbook.PSObject.BaseObject
    $raw[1]=$null
    $hasProfile = $addin.GetType('ExcelAddIn1.Funtion.WorkbookProjectProfileService').GetMethod('TryLoad').Invoke($null,$raw)
    Assert-V2 (-not $hasProfile) 'Preview unexpectedly mutated profile'
    $canceled = Invoke-V2 'WorkbookEstimateV2PackageService' 'Apply' @($workbook,$preview,$false,$null)
    Assert-V2 ([string]::IsNullOrEmpty($canceled)) 'Cancel mutated workbook'
    $backup = Invoke-V2 'WorkbookEstimateV2PackageService' 'Apply' @($workbook,$preview,$true,$null)
    Assert-V2 (Test-Path -LiteralPath $backup) 'Package backup missing'
    $profile = Invoke-V2 'WorkbookProjectProfileService' 'LoadRequired' @($workbook)
    Assert-V2 ($profile.RegulationPackageChecksum -eq $target.PackageChecksum) 'Pinned identity mismatch'
    $beforePayload = $profile | ForEach-Object { [ExcelAddIn1.Core.ProjectProfileSerializer]::Serialize($_) }
    $again = Invoke-V2 'WorkbookEstimateV2PackageService' 'Preview' @($workbook,$target,(Get-Date),(Get-Date))
    $fail = [Action]{ throw 'V2 injected metadata failure' }
    $rolledBack=$false
    try { Invoke-V2 'WorkbookEstimateV2PackageService' 'Apply' @($workbook,$again,$true,$fail) | Out-Null }
    catch { $rolledBack=$true }
    Assert-V2 $rolledBack 'Failure injection did not fail'
    $restored = Invoke-V2 'WorkbookProjectProfileService' 'LoadRequired' @($workbook)
    Assert-V2 ([ExcelAddIn1.Core.ProjectProfileSerializer]::Serialize($restored) -eq $beforePayload) 'Rollback profile mismatch'
    $workbook.Save()
    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    $workbook=$excel.Workbooks.Open($workingCopy,0,$false)
    $reopened=Invoke-V2 'WorkbookProjectProfileService' 'LoadRequired' @($workbook)
    Assert-V2 ($reopened.RegulationPackageChecksum -eq $target.PackageChecksum) 'Save/reopen lost package'
    $sheets = @(Invoke-V2 'WorkbookEstimateV2ReportService' 'ListSheets' @($workbook))
    $sheets | Format-Table CodeName,Name,PrintArea | Out-Host
    $printable = @($sheets | Where-Object { -not [string]::IsNullOrWhiteSpace($_.PrintArea) })
    Assert-V2 ($printable.Count -gt 0) 'Sample has no printable sheets'
    $selection = [string[]]@($printable[0].CodeName)
    $check = Invoke-V2 'WorkbookEstimateV2ReportService' 'Check' @($workbook,$selection)
    Assert-V2 $check.CanExport ('Report check failed: ' + ($check.Errors -join '; '))
    $pdf = Join-Path $runRoot 'V2-single.pdf'
    $sheetCount = $workbook.Worksheets.Count
    $activeName = $excel.ActiveSheet.Name
    Invoke-V2 'WorkbookEstimateV2ReportService' 'ExportPdf' @($workbook,$selection,$pdf,$true) | Out-Null
    Assert-V2 ((Get-Item $pdf).Length -gt 1000) 'PDF output empty'
    Assert-V2 ($workbook.Worksheets.Count -eq $sheetCount -and $excel.ActiveSheet.Name -eq $activeName) 'Export changed source selection/sheets'
    $testSheet = $workbook.Worksheets.Item($printable[0].Name)
    $oldName = $testSheet.Name
    $testSheet.Name = 'V2-Renamed'
    $renamed = Invoke-V2 'WorkbookEstimateV2ReportService' 'Check' @($workbook,$selection)
    Assert-V2 $renamed.CanExport 'CodeName did not survive rename'
    $testSheet.Name = $oldName
    $testArea = $testSheet.Range($testSheet.PageSetup.PrintArea)
    $testCell = $testArea.Cells.Item(1,1)
    $originalFormula = $testCell.Formula
    $testCell.Formula = '=1/0'
    $invalid = Invoke-V2 'WorkbookEstimateV2ReportService' 'Check' @($workbook,$selection)
    Assert-V2 (-not $invalid.CanExport) 'Formula error was not blocked'
    $pdfHash = Get-SharedHash $pdf
    $failedExport = $false
    try { Invoke-V2 'WorkbookEstimateV2ReportService' 'ExportPdf' @($workbook,$selection,$pdf,$true) | Out-Null }
    catch { $failedExport = $true }
    Assert-V2 ($failedExport -and (Get-SharedHash $pdf) -eq $pdfHash) 'Failed export replaced the existing PDF'
    $testCell.Formula = $originalFormula
    $oldPrintArea = $testSheet.PageSetup.PrintArea
    $testSheet.PageSetup.PrintArea = ''
    $noArea = Invoke-V2 'WorkbookEstimateV2ReportService' 'Check' @($workbook,$selection)
    Assert-V2 (-not $noArea.CanExport) 'Missing Print Area was not blocked'
    $testSheet.PageSetup.PrintArea = $oldPrintArea
    $metadataSheet = $workbook.Worksheets.Add()
    $metadataSheet.Name = 'V2-Print-QA'
    $metadataSheet.Cells.Item(1,1).Value2 = '__TTB_ROWTYPE'
    $metadataSheet.Range('A:A').EntireColumn.Hidden = $true
    $metadataSheet.PageSetup.PrintArea = '$A$1:$B$2'
    $metadataKey = if ([string]::IsNullOrWhiteSpace($metadataSheet.CodeName)) { 'name:' + $metadataSheet.Name } else { $metadataSheet.CodeName }
    $metadataCheck = Invoke-V2 'WorkbookEstimateV2ReportService' 'Check' @($workbook,[string[]]@($metadataKey))
    Assert-V2 (-not $metadataCheck.CanExport) 'Hidden technical metadata in Print Area was not blocked'
    $metadataSheet.Delete()
    if ($printable.Count -gt 1) {
        $multiple = [string[]]@($printable[0].CodeName,$printable[1].CodeName)
        Invoke-V2 'WorkbookEstimateV2ReportService' 'ExportPdf' @($workbook,$multiple,(Join-Path $runRoot 'V2-multiple.pdf'),$true) | Out-Null
    }
    $workbook.Save()
    Assert-V2 ((Get-SharedHash $WorkbookPath) -eq $sourceHash) 'Source workbook changed'
    'PASS V2-801 preview/cancel/pin/backup/injected rollback/save-reopen'
    'PASS V2-901 single/multiple PDF, rename identity, formula error blocking, source sheets/selection preserved'
    $compat = Invoke-V2 'WorkbookEstimateV2CompatibilityService' 'Prepare' @($workbook)
    $compat | Format-List | Out-Host
    $sources = @(Invoke-V2 'WorkbookEstimateV2RegistrationService' 'ListRegistered' @($workbook))
    Assert-V2 ($sources.Count -gt 0) 'Legacy sample was not registered'
    $state = Invoke-V2 'WorkbookEstimateV2StateService' 'LoadOrCreate' @($workbook)
    'V2 work items: ' + $state.WorkItems.Count
    $sources | Format-List | Out-Host
    $workbook.Save()
    $workbook.SaveCopyAs((Join-Path $runRoot 'HoaLuNamDinh-V2-user-test.xlsx'))
    # Explicit fixtures exercise grouping and formulas, not financial equivalence with the reference estimate.
    $fixture = $workbook.Worksheets.Add()
    $fixture.Name = 'V2-QA-Fixture'
    $headers = @('Ma cong tac','Dinh muc','Mo ta cong viec','DVT','Khoi luong')
    for ($c=1; $c -le 5; $c++) { $fixture.Cells.Item(1,$c).Value2 = $headers[$c-1] }
    for ($r=2; $r -le 4; $r++) {
        $fixture.Cells.Item($r,1).Value2 = 'QA-' + $r
        $fixture.Cells.Item($r,3).Value2 = 'Grouping fixture ' + $r
        $fixture.Cells.Item($r,4).Value2 = 'ha'
        $fixture.Cells.Item($r,5).Formula = '=1+1'
    }
    Invoke-V2 'WorkbookEstimateV2RegistrationService' 'RegisterSelectedRange' @($workbook,$fixture.Range('A1:E4')) | Out-Null
    $state = Invoke-V2 'WorkbookEstimateV2StateService' 'LoadOrCreate' @($workbook)
    $items = @($state.WorkItems | Where-Object { $_.SourceKey -like '*V2-QA-Fixture*' })
    if ($items.Count -ne 3) { $items = @($state.WorkItems | Where-Object { $_.SourceKey -like ('*' + $fixture.CodeName + '*') }) }
    Assert-V2 ($items.Count -eq 3) ('Fixture ID count incorrect: ' + ($state.WorkItems.SourceKey -join ';'))
    for ($i=0; $i -lt 3; $i++) {
        $variant = if ($i -eq 1) { 'forest-2' } else { 'forest-1' }
        Invoke-V2 'WorkbookEstimateV2StateService' 'BindNorm' @($workbook,$items[$i].WorkItemId,'NORM-000.0200',$variant,$target.PackageId,$target.DataVersion,$target.PackageChecksum) | Out-Null
    }
    $resourceWrite = Invoke-V2 'WorkbookEstimateV2ResourceSheetWriter' 'Apply' @($workbook)
    # Use writer metadata instead of assuming the input row.
    $resourceSheet = $workbook.Worksheets.Item($resourceWrite.WorksheetName)
    $resourceValues = $resourceSheet.UsedRange.Value2
    $metadataColumn = 0
    for ($c=1; $c -le $resourceValues.GetLength(1); $c++) {
        if ([string]$resourceValues[1,$c] -eq '__TTB_ROWTYPE') { $metadataColumn = $c; break }
    }
    Assert-V2 ($metadataColumn -gt 0) 'Resource metadata header not found'
    $priceFilled = $false
    for ($r=1; $r -le $resourceSheet.UsedRange.Rows.Count; $r++) {
        if ([string]$resourceValues[$r,$metadataColumn] -eq 'LABOR_DIRECT_INPUT') {
            $resourceSheet.Cells.Item($r,4).Value2 = 100000.0
            $priceFilled = $true
        }
    }
    if (-not $priceFilled) {
        $laborInputs = @{ COEFFICIENT=1.0; BASESALARY=2200000.0; WORKDAYS=22.0; DANGER=0.0; MOBILE=0.0 }
        foreach ($key in $laborInputs.Keys) {
            $name = [ExcelAddIn1.Core.EstimateV2ExcelNames]::LaborInput('LAB-QNCN-7',$key)
            $inputCell = $workbook.Names.Item($name).RefersToRange
            $resourceSheet.Cells.Item([int]$inputCell.Row,[int]$inputCell.Column).Value2 = [double]$laborInputs[$key]
        }
        $priceFilled=$true
    }
    $excel.CalculateFull()
    $rateWrite = Invoke-V2 'WorkbookEstimateV2RateSheetWriter' 'Apply' @($workbook,[ExcelAddIn1.Core.EstimateV2RateEnvironment]::Land,$null)
    Assert-V2 ($rateWrite.RateCount -eq 2) 'Same variant not deduplicated or different variant merged'
    Assert-V2 ($rateWrite.MissingPriceRateCount -eq 0) 'Labor fixture price missing'
    $costWrite = Invoke-V2 'WorkbookEstimateV2CostLinkWriter' 'Apply' @($workbook)
    $costWrite | Format-List | Out-Host
    Assert-V2 ($costWrite.LinkedWorkItemCount -eq 3) ('Cost link count incorrect: ' + $costWrite.LinkedWorkItemCount)
    $excel.CalculateFull()
    for ($r=2; $r -le 4; $r++) {
        Assert-V2 ($fixture.Cells.Item($r,5).Formula -eq '=1+1') 'Quantity formula overwritten'
        Assert-V2 ([string]$fixture.Cells.Item($r,7).Formula -like '=*') 'Unit labor price is not a formula'
    }
    $totalName = $workbook.Names.Item('TTBMVN_V2_GIADT_NC')
    Assert-V2 ([double]$totalName.RefersToRange.Value2 -eq 20900000.0) ('Fixture labor total wrong: ' + $totalName.RefersToRange.Value2)
    $payload = [ExcelAddIn1.Core.EstimateV2StateSerializer]::Serialize((Invoke-V2 'WorkbookEstimateV2StateService' 'LoadOrCreate' @($workbook)))
    Invoke-V2 'WorkbookEstimateV2ValidationService' 'ScanReadOnly' @($workbook) | Out-Null
    Assert-V2 ($payload -eq [ExcelAddIn1.Core.EstimateV2StateSerializer]::Serialize((Invoke-V2 'WorkbookEstimateV2StateService' 'LoadOrCreate' @($workbook)))) 'Read-only validation mutated legal state'
    $workbook.Save()
    'PASS migration, 3 bindings/2 variants/2 DG blocks, preserved quantity formulas, Excel NC total 20900000, read-only validation'
    'WorkingCopy: ' + $workingCopy
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    Stop-IsolatedExcelTestProcess $session
}
