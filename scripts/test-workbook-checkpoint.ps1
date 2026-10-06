param(
    [Parameter(Mandatory = $true)]
    [string]$CheckpointPath,
    [double]$ExpectedE27 = 1198731000,
    [int]$ExpectedFormulaErrors = 92
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$fullPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $CheckpointPath))
$checkpointRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau\Checkpoints"))
if (-not $fullPath.StartsWith(
    $checkpointRoot.TrimEnd('\') + '\',
    [StringComparison]::OrdinalIgnoreCase)) {
    throw "Checkpoint nam ngoai thu muc cho phep: $fullPath"
}
if (-not (Test-Path -LiteralPath $fullPath)) {
    throw "Khong tim thay checkpoint: $fullPath"
}

$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$profileServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$pinServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookRegulationPackageService", $true)
$roleServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$loadProfileMethod = $profileServiceType.GetMethod("LoadRequired")
$verifyPackageMethod = $pinServiceType.GetMethod("Verify")
$validateRolesMethod = $roleServiceType.GetMethod("Validate")
$defaultStore = Join-Path $env:LOCALAPPDATA "TTBMVNExcelTools\RegulationPackages"

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

$excelSession = Start-IsolatedExcelTestProcess
$excel = $excelSession.Application
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($fullPath, 0, $true)
    $profile = Invoke-Static $loadProfileMethod @($workbook)
    $roles = Invoke-Static $validateRolesMethod @($workbook)
    $package = Invoke-Static $verifyPackageMethod @($workbook, $defaultStore)
    if (-not $roles.IsValid) {
        throw "Worksheet role checkpoint khong hop le."
    }
    if ($package.Status.ToString() -ne "Available") {
        throw "Package pin checkpoint khong Available: $($package.Status)."
    }

    $summary = $workbook.Worksheets.Item("THKP-TC")
    try {
        $actualE27 = [double]$summary.Range("E27").Value2
    }
    finally {
        Release-ComObject $summary
    }
    if ($actualE27 -ne $ExpectedE27) {
        throw "THKP-TC!E27 sai. Expected=$ExpectedE27 Actual=$actualE27"
    }

    $formulaErrors = 0
    foreach ($sheet in $workbook.Worksheets) {
        $used = $null
        $errors = $null
        try {
            $used = $sheet.UsedRange
            try {
                $errors = $used.SpecialCells(-4123, 16)
                $formulaErrors += $errors.CountLarge
            }
            catch [Runtime.InteropServices.COMException] {
                # Excel throws when a sheet has no matching error cells.
            }
        }
        finally {
            Release-ComObject $errors
            Release-ComObject $used
            Release-ComObject $sheet
        }
    }
    if ($formulaErrors -ne $ExpectedFormulaErrors) {
        throw "So o loi cong thuc thay doi. Expected=$ExpectedFormulaErrors Actual=$formulaErrors"
    }

    [pscustomobject]@{
        Checkpoint = $fullPath
        ProfileSchema = $profile.SchemaVersion
        PackageId = $profile.RegulationPackageId
        PackageVersion = $profile.RegulationPackageVersion
        PackageChecksum = $profile.RegulationPackageChecksum
        PackageStatus = $package.Status
        Roles = "PASS"
        FormulaErrors = $formulaErrors
        E27 = $actualE27
    } | Format-List
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $excelSession
}

Get-FileHash -LiteralPath $fullPath -Algorithm SHA256 | Format-List Path,Hash
