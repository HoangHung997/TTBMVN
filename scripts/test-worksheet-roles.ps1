param(
    [string]$WorkbookPath,
    [switch]$ApplyBaselineMapping,
    [switch]$RenameProbe
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\00. Du toan TP 3.xlsm"
}

$expectedPath = [IO.Path]::GetFullPath($WorkbookPath)
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
if (-not (Test-Path $corePath) -or -not (Test-Path $addinPath)) {
    throw "Chua co Release build. Chay scripts\build-release.ps1 truoc."
}

$coreAssembly = [Reflection.Assembly]::LoadFrom($corePath)
$addinAssembly = [Reflection.Assembly]::LoadFrom($addinPath)
$roleType = $coreAssembly.GetType("ExcelAddIn1.Core.WorksheetRole", $true)
$serviceType = $addinAssembly.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$setRoleMethod = $serviceType.GetMethod("SetRole")
$validateMethod = $serviceType.GetMethod("Validate")
$resolveMethod = $serviceType.GetMethod("ResolveRequired")

$baselineMap = [ordered]@{
    "VL-NC-M" = "ResourcePrices"
    "DG Can" = "UnitRateLand"
    "DG Nuoc" = "UnitRateWater"
    "Gia DT TC" = "EstimateAppendix"
    "THKP-TC" = "CostSummary"
    "Tracuu" = "NormLookupView"
    "ChiPhi" = "CostRuleView"
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-StaticMethod([Reflection.MethodInfo]$Method, [object[]]$Arguments) {
    return $Method.Invoke($null, $Arguments)
}

$openSession = Get-OpenExcelWorkbookSession $expectedPath
if ($null -eq $openSession) {
    throw "Workbook chua mo trong Excel: $expectedPath"
}
$excel = $openSession.Application
$workbook = $openSession.Workbook
try {
    if ($ApplyBaselineMapping) {
        foreach ($entry in $baselineMap.GetEnumerator()) {
            $worksheet = $null
            try {
                $worksheet = $workbook.Worksheets.Item($entry.Key)
                $role = [Enum]::Parse($roleType, $entry.Value, $false)
                [void](Invoke-StaticMethod $setRoleMethod @($worksheet, $role))
            }
            finally {
                Release-ComObject $worksheet
            }
        }
    }

    $validation = Invoke-StaticMethod $validateMethod @($workbook)
    if (-not $validation.IsValid) {
        $messages = @($validation.Issues | ForEach-Object { $_.Message }) -join " | "
        throw "Worksheet role validation failed: $messages"
    }

    $resolvedRoles = @()
    foreach ($roleName in [Enum]::GetNames($roleType)) {
        $role = [Enum]::Parse($roleType, $roleName, $false)
        $worksheet = Invoke-StaticMethod $resolveMethod @($workbook, $role)
        try {
            $resolvedRoles += "$roleName=$($worksheet.Name)"
        }
        finally {
            Release-ComObject $worksheet
        }
    }

    if ($RenameProbe) {
        $landRole = [Enum]::Parse($roleType, "UnitRateLand", $false)
        $worksheet = Invoke-StaticMethod $resolveMethod @($workbook, $landRole)
        $oldName = $worksheet.Name
        try {
            $probeName = "TTBMVN_ROLE_PROBE"
            $worksheet.Name = $probeName
            Release-ComObject $worksheet
            $worksheet = Invoke-StaticMethod $resolveMethod @($workbook, $landRole)
            if ($worksheet.Name -ne $probeName) {
                throw "Resolver khong theo duoc sheet sau rename."
            }
            $worksheet.Name = $oldName
        }
        finally {
            if ($null -ne $worksheet) {
                try {
                    if ($worksheet.Name -eq "TTBMVN_ROLE_PROBE") {
                        $worksheet.Name = $oldName
                    }
                }
                catch {
                }
                Release-ComObject $worksheet
            }
        }
    }

    if ($ApplyBaselineMapping -or $RenameProbe) {
        $workbook.Save()
    }

    [pscustomobject]@{
        Workbook = $workbook.FullName
        Validation = "PASS"
        RenameProbe = if ($RenameProbe) { "PASS" } else { "SKIP" }
        Roles = $resolvedRoles -join "; "
        Saved = $workbook.Saved
    } | Format-List
}
finally {
    Release-OpenExcelWorkbookSession $openSession
}
