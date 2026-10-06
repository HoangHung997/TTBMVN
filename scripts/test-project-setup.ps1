param(
    [string]$WorkbookPath,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\00. Du toan TP 3.xlsm"
}
$sourcePath = [IO.Path]::GetFullPath($WorkbookPath)
$variantRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau\Variants"))
$workingCopy = Join-Path $variantRoot "DT-401-working-copy.xlsm"
$variantPrefix = $variantRoot.TrimEnd('\') + '\'
if (-not $workingCopy.StartsWith($variantPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Working copy nam ngoai Variants: $workingCopy"
}
if (Test-Path -LiteralPath $workingCopy) {
    if (-not $Overwrite) {
        throw "Working copy da ton tai: $workingCopy. Dung -Overwrite de tao lai."
    }
    Remove-Item -LiteralPath $workingCopy -Force
}
New-Item -ItemType Directory -Path $variantRoot -Force | Out-Null
Copy-Item -LiteralPath $sourcePath -Destination $workingCopy

$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
$core = [Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
if (-not ("TTBMVN.Tests.ProjectSetupFaultSupport" -as [type])) {
    Add-Type -ReferencedAssemblies @($corePath, $addinPath) -TypeDefinition @"
using System;
using System.IO;
using ExcelAddIn1.Funtion;

namespace TTBMVN.Tests
{
    public static class ProjectSetupFaultSupport
    {
        public static Action<ProjectSetupCommitPhase> CreateAfterProfileFault()
        {
            return phase =>
            {
                if (phase == ProjectSetupCommitPhase.ProfileSaved)
                    throw new IOException("Injected DT-401 failure after profile save.");
            };
        }
    }
}
"@
}

$profileService = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$roleService = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$setupService = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectSetupService", $true)
$bootstrapService = $addin.GetType("ExcelAddIn1.Funtion.RegulationPackageBootstrapService", $true)
$sheetCoordinator = $addin.GetType("ExcelAddIn1.Funtion.WorkbookSheetChangeCoordinator", $true)
$formType = $addin.GetType("ExcelAddIn1.Winform.FrmProjectSetup", $true)
$draftType = $core.GetType("ExcelAddIn1.Core.ProjectSetupDraft", $true)
$plannerType = $core.GetType("ExcelAddIn1.Core.ProjectSetupPlanner", $true)
$mappingBuilder = $core.GetType("ExcelAddIn1.Core.WorksheetRoleMappingBuilder", $true)
$transitionCatalog = $core.GetType("ExcelAddIn1.Core.RegulationTransitionRuleCatalog", $true)

$readPayloadMethod = $profileService.GetMethod("TryReadPayload")
$loadProfileMethod = $profileService.GetMethod("LoadRequired")
$readAssignmentsMethod = $roleService.GetMethod("ReadAssignments")
$clearAssignmentsMethod = $roleService.GetMethod("ClearAssignments")
$validateRolesMethod = $roleService.GetMethod("Validate")
$captureSheetsMethod = $sheetCoordinator.GetMethod("CaptureSnapshot")
$buildMappingMethod = $mappingBuilder.GetMethod("FromAssignments")
$createPlanMethod = $plannerType.GetMethod("CreatePlan")
$loadPackagesMethod = $bootstrapService.GetMethod("LoadAvailablePackages")
$commitMethod = $setupService.GetMethods() | Where-Object {
    $_.Name -eq "Commit" -and $_.GetParameters().Length -eq 2
} | Select-Object -First 1
$commitWithCallbackMethod = $setupService.GetMethods() | Where-Object {
    $_.Name -eq "Commit" -and $_.GetParameters().Length -eq 3
} | Select-Object -First 1

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
    $result = $Method.Invoke($null, $invokeArguments)
    Write-Output -NoEnumerate $result
}

function Get-ProfilePayload([object]$Workbook) {
    $arguments = New-Object "object[]" 2
    $arguments[0] = $Workbook.PSObject.BaseObject
    $arguments[1] = $null
    if (-not [bool]$readPayloadMethod.Invoke($null, $arguments)) {
        return $null
    }
    return [string]$arguments[1]
}

function Get-AssignmentSignature([object]$Workbook) {
    $assignments = Invoke-Static $readAssignmentsMethod @($Workbook)
    return (($assignments | ForEach-Object {
        $_.SheetKey + "|" + $_.SheetName + "|" + $_.RoleId
    } | Sort-Object) -join "`n")
}

function New-ProjectSetupPlan(
    [object]$Workbook,
    [object]$Packages,
    [DateTime]$PreparedDate,
    [Nullable[DateTime]]$ApprovalDate,
    [DateTime]$EvaluationDate,
    [DateTime]$PriceDate) {
    $workbookArguments = New-Object "object[]" 1
    $workbookArguments[0] = $Workbook.PSObject.BaseObject
    $sheets = $captureSheetsMethod.Invoke($null, $workbookArguments)
    $assignments = $readAssignmentsMethod.Invoke($null, $workbookArguments)
    $mappingArguments = New-Object "object[]" 2
    $mappingArguments[0] = $assignments.PSObject.BaseObject
    $mappingArguments[1] = $sheets.PSObject.BaseObject
    $mapping = $buildMappingMethod.Invoke($null, $mappingArguments)
    $draftArguments = New-Object "object[]" 8
    $draftArguments[0] = "DT-401-INTEGRATION"
    $draftArguments[1] = $PreparedDate.Date
    $draftArguments[2] = $ApprovalDate
    $draftArguments[3] = $EvaluationDate.Date
    $draftArguments[4] = $PriceDate.Date
    $draftArguments[5] = "PRICE-DT401"
    $draftArguments[6] = ""
    $draftArguments[7] = $mapping.PSObject.BaseObject
    $draft = [Activator]::CreateInstance($draftType, $draftArguments)
    $rules = $transitionCatalog.GetProperty("All").GetValue($null)
    $planningArguments = New-Object "object[]" 4
    $planningArguments[0] = $draft.PSObject.BaseObject
    $planningArguments[1] = $sheets.PSObject.BaseObject
    $planningArguments[2] = $Packages.PSObject.BaseObject
    $planningArguments[3] = $rules
    $planning = $createPlanMethod.Invoke($null, $planningArguments)
    if (-not $planning.IsValid) {
        throw "Project setup plan khong hop le: $($planning.Errors -join ' ')"
    }
    return $planning.Plan
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
                    finally {
                        Release-ComObject $errors
                    }
                }
            }
            finally {
                Release-ComObject $worksheet
            }
        }
    }
    finally {
        Release-ComObject $worksheets
    }
    return $total
}

$packages = $loadPackagesMethod.Invoke($null, (New-Object "object[]" 0))
if (($packages | Where-Object PackageId -eq "BQP-RPBM-2021").Count -lt 1 -or
    ($packages | Where-Object {
        $_.PackageId -eq "BQP-RPBM-2025" -and $_.DataVersion -eq "2.0.1"
    }).Count -lt 1) {
    throw "Kho package thieu BQP-RPBM-2021 hoac ban va 2025@2.0.1."
}

$excelSession = Start-IsolatedExcelTestProcess
$excel = $excelSession.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
$verifyBook = $null
try {
    $initialPayload = Get-ProfilePayload $workbook
    $initialRoles = Get-AssignmentSignature $workbook
    $initialSaved = $workbook.Saved
    $initialErrorCount = Get-ExcelErrorCount $workbook
    $oldScreenUpdating = $excel.ScreenUpdating
    $oldCalculation = $excel.Calculation
    $oldEnableEvents = $excel.EnableEvents
    $oldDisplayAlerts = $excel.DisplayAlerts

    # UI smoke: construction and preview must be valid without writing workbook data.
    $formArguments = New-Object "object[]" 2
    $formArguments[0] = $workbook.PSObject.BaseObject
    $formArguments[1] = $null
    $setupForm = [Activator]::CreateInstance($formType, $formArguments)
    try {
        $saveField = $formType.GetField(
            "saveButton",
            [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
        $statusField = $formType.GetField(
            "statusLabel",
            [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
        $titleMatches = -not [string]::IsNullOrWhiteSpace([string]$setupForm.Text)
        $buttonEnabled = [bool]$saveField.GetValue($setupForm).Enabled
        if (-not $titleMatches -or -not $buttonEnabled) {
            throw ("Project setup form khong khoi tao duoc preview hop le. " +
                "Title='$($setupForm.Text)'; TitleMatches=$titleMatches; Enabled=$buttonEnabled; " +
                "Status='$($statusField.GetValue($setupForm).Text)'.")
        }
    }
    finally {
        $setupForm.Dispose()
    }

    # Cancel path: planning is pure and must not change the workbook.
    $plan = New-ProjectSetupPlan $workbook $packages `
        ([DateTime]::new(2026, 8, 1)) $null `
        ([DateTime]::new(2026, 8, 3)) ([DateTime]::new(2026, 8, 1))
    if ((Get-ProfilePayload $workbook) -ne $initialPayload -or
        (Get-AssignmentSignature $workbook) -ne $initialRoles -or
        $workbook.Saved -ne $initialSaved) {
        throw "Cancel/planning path da thay doi workbook."
    }

    # Inject failure after profile save and verify exact rollback.
    $fault = [TTBMVN.Tests.ProjectSetupFaultSupport]::CreateAfterProfileFault()
    $faultCaught = $false
    try {
        Invoke-Static $commitWithCallbackMethod @($workbook, $plan, $fault) | Out-Null
    }
    catch {
        $faultCaught = $true
    }
    if (-not $faultCaught -or
        (Get-ProfilePayload $workbook) -ne $initialPayload -or
        (Get-AssignmentSignature $workbook) -ne $initialRoles) {
        throw "Fault path khong rollback dung profile/mapping."
    }

    # Untagged workbook: clear role tags, rebuild from canonical names, then commit.
    Invoke-Static $clearAssignmentsMethod @($workbook) | Out-Null
    if ((Invoke-Static $validateRolesMethod @($workbook)).IsValid) {
        throw "Fixture untagged van con role hop le."
    }
    $untaggedPlan = New-ProjectSetupPlan $workbook $packages `
        ([DateTime]::new(2026, 8, 1)) $null `
        ([DateTime]::new(2026, 8, 3)) ([DateTime]::new(2026, 8, 1))
    Invoke-Static $commitMethod @($workbook, $untaggedPlan) | Out-Null
    if (-not (Invoke-Static $validateRolesMethod @($workbook)).IsValid) {
        throw "Untagged workbook khong duoc mapping lai."
    }

    # Rename all seven business sheets. CodeName keeps the mapping stable.
    $renameByRole = @{
        ResourcePrices = "DT401 01 Gia nguon"
        UnitRateLand = "DT401 02 Don gia can"
        UnitRateWater = "DT401 03 Don gia nuoc"
        EstimateAppendix = "DT401 04 Phu luc"
        CostSummary = "DT401 05 Tong hop"
        NormLookupView = "DT401 06 Dinh muc"
        CostRuleView = "DT401 07 Chi phi"
    }
    $assignments = Invoke-Static $readAssignmentsMethod @($workbook)
    foreach ($assignment in $assignments) {
        $sheet = $workbook.Worksheets.Item($assignment.SheetName)
        try {
            $sheet.Name = $renameByRole[$assignment.RoleId]
        }
        finally {
            Release-ComObject $sheet
        }
    }
    $renamedPlan = New-ProjectSetupPlan $workbook $packages `
        ([DateTime]::new(2026, 8, 1)) $null `
        ([DateTime]::new(2026, 8, 3)) ([DateTime]::new(2026, 8, 1))
    Invoke-Static $commitMethod @($workbook, $renamedPlan) | Out-Null
    if (-not (Invoke-Static $validateRolesMethod @($workbook)).IsValid) {
        throw "Mapping theo CodeName hong sau khi doi ten sheet."
    }

    $profile = Invoke-Static $loadProfileMethod @($workbook)
    if ($profile.RegulationPackageId -ne "BQP-RPBM-2025" -or
        $profile.RegulationPackageVersion -ne "2.0.1" -or
        $profile.PriceProfileId -ne "PRICE-DT401") {
        throw "Profile sau commit khong dung package/moc gia."
    }
    $workbook.Save()
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null

    $verifyBook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $reopenedProfile = Invoke-Static $loadProfileMethod @($verifyBook)
    $reopenedRoles = Invoke-Static $validateRolesMethod @($verifyBook)
    $finalErrorCount = Get-ExcelErrorCount $verifyBook
    if ($reopenedProfile.RegulationPackageId -ne "BQP-RPBM-2025" -or
        -not $reopenedRoles.IsValid -or
        $finalErrorCount -ne $initialErrorCount) {
        throw "Save/reopen regression: profile, roles hoac Excel errors thay doi."
    }
    $summary = $verifyBook.Worksheets.Item("DT401 05 Tong hop")
    try {
        $summaryE27 = [double]$summary.Range("E27").Value2
        if ($summaryE27 -ne 1198731000) {
            throw "Baseline E27 regression."
        }
    }
    finally {
        Release-ComObject $summary
    }

    if ($excel.ScreenUpdating -ne $oldScreenUpdating -or
        $excel.Calculation -ne $oldCalculation -or
        $excel.EnableEvents -ne $oldEnableEvents -or
        $excel.DisplayAlerts -ne $oldDisplayAlerts) {
        throw "Excel application state khong duoc khoi phuc."
    }

    [pscustomobject]@{
        Workbook = $workingCopy
        Packages = (($packages | ForEach-Object {
            $_.PackageId + "@" + $_.DataVersion
        }) -join "; ")
        CancelNoWrite = $true
        FormSmoke = "PASS"
        FaultRollback = $true
        UntaggedMapping = "PASS"
        RenamedSheetMapping = "PASS"
        SaveReopen = "PASS"
        ExcelErrors = $finalErrorCount
        SummaryE27 = $summaryE27
    } | Format-List
}
finally {
    if ($null -ne $verifyBook) {
        try { $verifyBook.Close($false) } catch {}
        Release-ComObject $verifyBook
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    Stop-IsolatedExcelTestProcess $excelSession
}
