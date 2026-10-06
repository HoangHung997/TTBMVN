param(
    [string]$WorkbookPath,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-504-start.xlsm"
}
$sourcePath = [IO.Path]::GetFullPath($WorkbookPath)
$variantRoot = Join-Path $repoRoot "Dutoanmau\Variants"
$checkpointRoot = Join-Path $repoRoot "Dutoanmau\Checkpoints"
$workingCopy = Join-Path $variantRoot "DT-504-working-copy.xlsm"
$restoreCopy = Join-Path $variantRoot "DT-504-restored-backup.xlsm"
$applyBackup = Join-Path $checkpointRoot "DT-504-apply-backup.xlsm"
$faultPhases = @(
    "BackupCreated",
    "ProfilePinned",
    "EstimateWritten",
    "CostSummaryWritten",
    "ValidationCompleted",
    "WorkbookSaved"
)
$faultBackups = @{}
foreach ($phase in $faultPhases) {
    $faultBackups[$phase] = Join-Path $checkpointRoot ("DT-504-fault-" + $phase + ".xlsm")
}
$cancelBackup = Join-Path $checkpointRoot "DT-504-cancel-should-not-exist.xlsm"
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
$storeRoot = Join-Path $env:LOCALAPPDATA "TTBMVNExcelTools\RegulationPackages"

New-Item -ItemType Directory -Path $variantRoot,$checkpointRoot -Force | Out-Null
$generated = @($workingCopy, $restoreCopy, $applyBackup, $cancelBackup) + @($faultBackups.Values)
foreach ($path in $generated) {
    if (Test-Path -LiteralPath $path) {
        if (-not $Overwrite) {
            throw "File test da ton tai: $path. Dung -Overwrite de tao lai."
        }
        Remove-Item -LiteralPath $path -Force
    }
}
Copy-Item -LiteralPath $sourcePath -Destination $workingCopy

$core = [Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
if (-not ("TTBMVN.Tests.MigrationResultFaultSupport" -as [type])) {
    Add-Type -ReferencedAssemblies @($corePath, $addinPath) -TypeDefinition @"
using System;
using System.IO;
using ExcelAddIn1.Funtion;

namespace TTBMVN.Tests
{
    public static class MigrationResultFaultSupport
    {
        public static Action<WorkbookPackageMigrationPhase> Create(WorkbookPackageMigrationPhase target)
        {
            return phase =>
            {
                if (phase == target)
                    throw new IOException("Injected DT-504 failure at " + target + ".");
            };
        }
    }
}
"@
}

$bootstrapType = $addin.GetType("ExcelAddIn1.Funtion.RegulationPackageBootstrapService", $true)
$pinType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookRegulationPackageService", $true)
$profileType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$priceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPriceProfileService", $true)
$auditType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookResultAuditService", $true)
$unitType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookUnitRateService", $true)
$estimateServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookEstimateAppendixService", $true)
$estimateWriterType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookEstimateAppendixWriter", $true)
$costServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookCostSummaryService", $true)
$costWriterType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookCostSummaryWriter", $true)
$migrationType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPackageMigrationService", $true)
$snapshotType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPackageMigrationSnapshot", $true)
$validationType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookValidationService", $true)

$loadPackages = $bootstrapType.GetMethod("LoadAvailablePackages")
$pin = $pinType.GetMethod("PinAndSave", [type[]]@(
    [Microsoft.Office.Interop.Excel.Workbook],
    [ExcelAddIn1.Core.RegulationPackage]))
$loadProfile = $profileType.GetMethod("LoadRequired")
$loadUnit = $unitType.GetMethod("LoadRequired")
$restorePlan = $auditType.GetMethod(
    "TryRestoreEstimatePlanForValidation",
    [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
$readConditions = $auditType.GetMethod("ReadEstimateConditions")
$loadAudit = $auditType.GetMethod("Load")
$estimatePreview = $estimateServiceType.GetMethod("Preview")
$writeEstimate = $estimateWriterType.GetMethod("Apply")
$loadCost = $costServiceType.GetMethod("Load")
$costPreview = $costServiceType.GetMethod("Preview")
$writeCost = $costWriterType.GetMethod("Apply")
$migrationPreview = $migrationType.GetMethod("Preview")
$migrationApply = $migrationType.GetMethods() | Where-Object {
    $_.Name -eq "Apply" -and $_.GetParameters().Length -eq 5
} | Select-Object -First 1
$migrationApplyFault = $migrationType.GetMethods() | Where-Object {
    $_.Name -eq "Apply" -and $_.GetParameters().Length -eq 6
} | Select-Object -First 1
$captureSnapshot = $snapshotType.GetMethod(
    "Capture",
    [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
$snapshotFingerprint = $snapshotType.GetProperty(
    "Fingerprint",
    [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
$scan = $validationType.GetMethod("Scan")

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-Static([Reflection.MethodInfo]$Method, [object[]]$Arguments) {
    $invokeArguments = New-Object "object[]" $Arguments.Length
    for ($index = 0; $index -lt $Arguments.Length; $index++) {
        $invokeArguments[$index] = if ($null -eq $Arguments[$index]) {
            $null
        }
        else {
            $Arguments[$index].PSObject.BaseObject
        }
    }
    return $Method.Invoke($null, $invokeArguments)
}

function Get-Package([object[]]$Packages, [string]$Id) {
    $package = $Packages |
        Where-Object PackageId -eq $Id |
        Sort-Object { [version]$_.DataVersion } -Descending |
        Select-Object -First 1
    if ($null -eq $package) { throw "Khong co package $Id." }
    return $package
}

function Restore-Plan([object]$Workbook, [object]$Worksheet, [object]$Context) {
    $arguments = New-Object "object[]" 4
    $arguments[0] = $Workbook.PSObject.BaseObject
    $arguments[1] = $Worksheet.PSObject.BaseObject
    $arguments[2] = $Context.PSObject.BaseObject
    $arguments[3] = $null
    if (-not [bool]$restorePlan.Invoke($null, $arguments)) {
        throw "Khong khoi phuc duoc plan phu luc tu audit."
    }
    return $arguments[3]
}

function Invoke-EstimatePreview([object]$Context, [object]$Plan, [string[]]$Conditions) {
    return Invoke-Static $estimatePreview @($Context, $Plan, $Conditions)
}

function Get-StateFingerprint([object]$Workbook, [object]$Plan) {
    $snapshot = Invoke-Static $captureSnapshot @($Workbook, $Plan)
    return [string]$snapshotFingerprint.GetValue($snapshot, $null)
}

function Assert-CleanValidation([object]$Workbook, [string]$Stage) {
    $report = Invoke-Static $scan @($Workbook)
    if (-not $report.IsValid) {
        $errors = $report.Issues | Where-Object IsBlocking | ForEach-Object { "$($_.Code): $($_.Message)" }
        throw "$Stage validation loi: $($errors -join ' | ')"
    }
    return $report
}

function Get-ValidationSignature([object]$Workbook) {
    $report = Invoke-Static $scan @($Workbook)
    $items = @($report.Issues | ForEach-Object {
        "$($_.Severity)|$($_.Code)|$($_.WorksheetRoleId)|$($_.WorksheetCodeName)|$($_.Address)|$($_.Subject)"
    } | Sort-Object)
    return "$($report.ErrorCount)/$($report.WarningCount):$($items -join ';')"
}

function Get-ExcelErrorCount([object]$Workbook) {
    [long]$total = 0
    $worksheets = $Workbook.Worksheets
    try {
        for ($index = 1; $index -le $worksheets.Count; $index++) {
            $worksheet = $worksheets.Item($index)
            try {
                foreach ($cellType in @(-4123, 2)) {
                    $errors = $null
                    try {
                        $errors = $worksheet.Cells.SpecialCells($cellType, 16)
                        $total += [long]$errors.CountLarge
                    }
                    catch [Runtime.InteropServices.COMException] {
                    }
                    finally { Release-ComObject $errors }
                }
            }
            finally { Release-ComObject $worksheet }
        }
    }
    finally { Release-ComObject $worksheets }
    return $total
}

$packages = @($loadPackages.Invoke($null, (New-Object "object[]" 0)))
$package2021 = Get-Package $packages "BQP-RPBM-2021"
$package2025 = Get-Package $packages "BQP-RPBM-2025"

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$verifyBook = $null
$backupBook = $null
$restoreBook = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $initialErrors = Get-ExcelErrorCount $workbook

    # Keep the existing legacy/result cells as the 2021 source baseline. Package 2021
    # intentionally has no detailed norm consumption table for the new calculator.
    $targetContextBeforeMigration = Invoke-Static $loadUnit @($workbook)
    $estimateSheet = $workbook.Worksheets.Item("Gia DT TC")
    try {
        $sourcePlan = Restore-Plan $workbook $estimateSheet $targetContextBeforeMigration
        $conditions = @((Invoke-Static $readConditions @($workbook, $estimateSheet.CodeName)))
    }
    finally { Release-ComObject $estimateSheet }
    $sourceSummary = $workbook.Worksheets.Item("THKP-TC")
    try { $sourceTotal = [decimal]$sourceSummary.Range("E27").Value2 }
    finally { Release-ComObject $sourceSummary }
    Invoke-Static $pin @($workbook, $package2021) | Out-Null
    $workbook.Save()

    $sourceValidationSignature = Get-ValidationSignature $workbook
    $sourceFingerprint = Get-StateFingerprint $workbook $sourcePlan
    $sourceAudit = Invoke-Static $loadAudit @($workbook)
    $sourceProfile = Invoke-Static $loadProfile @($workbook)

    $preview = Invoke-Static $migrationPreview @(
        $workbook,
        $storeRoot,
        $package2025.PackageId,
        $package2025.DataVersion,
        $package2025.PackageChecksum)
    if (-not $preview.Impact.CanApply -or
        $preview.Impact.EstimateLineCount -ne 13 -or
        $preview.Impact.AuditEntryCount -ne 22 -or
        $preview.Impact.ValueChanges.Count -ne 7) {
        throw "Preview DT-504 khong day du hoac con loi: $($preview.Impact.Issues -join ' | ')"
    }
    $targetTotal = [decimal](
        $preview.Impact.ValueChanges |
        Where-Object Key -eq "CostSummary.RoundedAfterTax" |
        Select-Object -ExpandProperty TargetValue)
    $reportedSourceTotal = [decimal](
        $preview.Impact.ValueChanges |
        Where-Object Key -eq "CostSummary.RoundedAfterTax" |
        Select-Object -ExpandProperty SourceValue)
    if ($reportedSourceTotal -ne $sourceTotal) {
        throw "Migration report khong lay dung gia tri nguon trong workbook."
    }

    $canceled = Invoke-Static $migrationApply @(
        $workbook, $storeRoot, $preview, $cancelBackup, $false)
    if ($canceled.Status.ToString() -ne "Canceled" -or
        (Test-Path $cancelBackup) -or
        (Get-StateFingerprint $workbook $sourcePlan) -ne $sourceFingerprint) {
        throw "Cancel migration da thay doi workbook."
    }

    $faultResults = @()
    foreach ($phaseName in $faultPhases) {
        $phase = [ExcelAddIn1.Funtion.WorkbookPackageMigrationPhase]::$phaseName
        $callback = [TTBMVN.Tests.MigrationResultFaultSupport]::Create($phase)
        $caught = $false
        try {
            Invoke-Static $migrationApplyFault @(
                $workbook,
                $storeRoot,
                $preview,
                $faultBackups[$phaseName],
                $true,
                $callback) | Out-Null
        }
        catch {
            $caught = $true
            $errorObject = $_.Exception
            while ($null -ne $errorObject.InnerException -and
                $errorObject.GetType().FullName -ne "ExcelAddIn1.Funtion.WorkbookPackageMigrationException") {
                $errorObject = $errorObject.InnerException
            }
            if ($phaseName -ne "BackupCreated" -and
                (-not $errorObject.RollbackAttempted -or -not $errorObject.RollbackSucceeded)) {
                throw "Fault $phaseName khong rollback thanh cong."
            }
        }
        if (-not $caught -or -not (Test-Path $faultBackups[$phaseName])) {
            throw "Fault $phaseName khong tao exception/backup."
        }
        $actualFingerprint = Get-StateFingerprint $workbook $sourcePlan
        if ($actualFingerprint -ne $sourceFingerprint) {
            throw "Fault $phaseName khong khoi phuc dung fingerprint source."
        }
        $faultProfile = Invoke-Static $loadProfile @($workbook)
        if ($faultProfile.RegulationPackageId -ne "BQP-RPBM-2021") {
            throw "Fault $phaseName khong khoi phuc package 2021."
        }
        $faultResults += $phaseName
    }
    if ((Get-ValidationSignature $workbook) -ne $sourceValidationSignature) {
        throw "Ma tran fault khong khoi phuc dung validation signature source."
    }

    $applied = Invoke-Static $migrationApply @(
        $workbook, $storeRoot, $preview, $applyBackup, $true)
    if ($applied.Status.ToString() -ne "Applied" -or
        -not $applied.ValidationReport.IsValid -or
        -not (Test-Path $applyBackup)) {
        throw "Apply migration 2025 khong thanh cong."
    }
    $workbook.Save()
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null

    $verifyBook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $finalProfile = Invoke-Static $loadProfile @($verifyBook)
    $finalAudit = Invoke-Static $loadAudit @($verifyBook)
    $finalValidation = Assert-CleanValidation $verifyBook "Target 2025 reopen"
    $summary = $verifyBook.Worksheets.Item("THKP-TC")
    try { $finalE27 = [decimal]$summary.Range("E27").Value2 }
    finally { Release-ComObject $summary }
    if ($finalProfile.RegulationPackageId -ne "BQP-RPBM-2025" -or
        $finalProfile.RegulationPackageVersion -ne "2.0.1" -or
        $finalAudit.Entries.Count -ne 22 -or
        $finalE27 -ne $targetTotal) {
        throw "Target sau reopen khong khop migration report."
    }
    $finalErrors = Get-ExcelErrorCount $verifyBook
    if ($finalErrors -ne $initialErrors) {
        throw "So o loi Excel thay doi: before=$initialErrors after=$finalErrors."
    }
    $verifyBook.Close($false)
    Release-ComObject $verifyBook
    $verifyBook = $null

    # The apply backup is a restorable 2021 checkpoint, independent of in-memory rollback.
    Copy-Item -LiteralPath $applyBackup -Destination $restoreCopy
    $restoreBook = $excel.Workbooks.Open($restoreCopy, 0, $true)
    $restoredProfile = Invoke-Static $loadProfile @($restoreBook)
    $restoredAudit = Invoke-Static $loadAudit @($restoreBook)
    $restoredSummary = $restoreBook.Worksheets.Item("THKP-TC")
    try { $restoredE27 = [decimal]$restoredSummary.Range("E27").Value2 }
    finally { Release-ComObject $restoredSummary }
    if ($restoredProfile.RegulationPackageId -ne "BQP-RPBM-2021" -or
        $restoredAudit.Entries.Count -ne $sourceAudit.Entries.Count -or
        $restoredE27 -ne $sourceTotal) {
        throw "Restore tu checkpoint backup khong ve dung source 2021."
    }
    if ((Get-ValidationSignature $restoreBook) -ne $sourceValidationSignature) {
        throw "Restore checkpoint khong khop validation signature source."
    }

    [pscustomobject]@{
        Workbook = $workingCopy
        SourcePackage = "$($sourceProfile.RegulationPackageId)@$($sourceProfile.RegulationPackageVersion)"
        TargetPackage = "$($finalProfile.RegulationPackageId)@$($finalProfile.RegulationPackageVersion)"
        PlanId = $preview.Plan.PlanId
        PackageChanges = $preview.Plan.Diff.Changes.Count
        CalculationChanges = $preview.Plan.Diff.CalculationDataChangeCount
        EstimateLines = $preview.Impact.EstimateLineCount
        AuditEntries = $finalAudit.Entries.Count
        PriceProfile = "$($preview.Impact.PriceProfileId)@$($preview.Impact.PriceProfileVersion)"
        SourceTotal = $sourceTotal
        TargetTotal = $targetTotal
        Difference = $targetTotal - $sourceTotal
        Cancel = "PASS"
        FaultPhases = $faultResults -join ","
        RestoreCheckpoint = "PASS"
        Validation = "PASS"
        ExcelErrors = $finalErrors
        ApplyBackup = $applyBackup
    } | Format-List
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $verifyBook) { $verifyBook.Close($false) }
    if ($null -ne $backupBook) { $backupBook.Close($false) }
    if ($null -ne $restoreBook) { $restoreBook.Close($false) }
    Release-ComObject $restoreBook
    Release-ComObject $backupBook
    Release-ComObject $verifyBook
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}

Get-FileHash -LiteralPath $workingCopy,$applyBackup,$restoreCopy -Algorithm SHA256 |
    Format-Table -AutoSize
