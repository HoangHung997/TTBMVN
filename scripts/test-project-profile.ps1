param(
    [string]$WorkbookPath,
    [switch]$ApplyTestProfile,
    [switch]$Reopen
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\00. Du toan TP 3.xlsm"
}
$expectedPath = [IO.Path]::GetFullPath($WorkbookPath)

$coreAssembly = [Reflection.Assembly]::LoadFrom(
    (Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"))
$addinAssembly = [Reflection.Assembly]::LoadFrom(
    (Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"))
$profileType = $coreAssembly.GetType("ExcelAddIn1.Core.ProjectProfile", $true)
$serializerType = $coreAssembly.GetType("ExcelAddIn1.Core.ProjectProfileSerializer", $true)
$profileServiceType = $addinAssembly.GetType(
    "ExcelAddIn1.Funtion.WorkbookProjectProfileService",
    $true)
$roleServiceType = $addinAssembly.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)

$saveMethod = $profileServiceType.GetMethod("Save")
$loadMethod = $profileServiceType.GetMethod("LoadRequired")
$readPayloadMethod = $profileServiceType.GetMethod("TryReadPayload")
$checksumMethod = $serializerType.GetMethod("ComputeChecksum")
$validateRolesMethod = $roleServiceType.GetMethod("Validate")

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-Static([Reflection.MethodInfo]$Method, [object[]]$Arguments) {
    return $Method.Invoke($null, $Arguments)
}

function Find-Workbook([object]$Excel, [string]$FullPath) {
    $workbooks = $Excel.Workbooks
    try {
        for ($index = 1; $index -le $workbooks.Count; $index++) {
            $candidate = $workbooks.Item($index)
            if ([string]::Equals($candidate.FullName, $FullPath, [StringComparison]::OrdinalIgnoreCase)) {
                return $candidate
            }
            Release-ComObject $candidate
        }
    }
    finally {
        Release-ComObject $workbooks
    }
    return $null
}

$openSession = Get-OpenExcelWorkbookSession $expectedPath
if ($null -eq $openSession) {
    throw "Workbook chua mo trong Excel: $expectedPath"
}
$excel = $openSession.Application
$workbook = $openSession.Workbook
try {
    $firstChanged = $false
    $secondChanged = $false
    if ($ApplyTestProfile) {
        $profile = [Activator]::CreateInstance($profileType)
        $profile.ProjectId = "TEST-DT-103-00-DU-TOAN-TP3"
        $profile.PreparedDate = [DateTime]::new(2026, 8, 3)
        $profile.ApprovalDate = $null
        $profile.PriceDate = [DateTime]::new(2026, 8, 3)
        $profile.RegulationPackageId = "TEST-DT103-PACKAGE-NOT-FOR-CALCULATION"
        $profile.RegulationPackageVersion = "1.0.0"
        $profile.RegulationPackageChecksum = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
        $profile.PriceProfileId = "TEST-DT103-PRICE-BASELINE"
        $profile.OverrideSummary = (
            "DT-103 persistence profile only; no business override. " * 20).Trim()

        $firstChanged = [bool](Invoke-Static $saveMethod @($workbook, $profile))
        $secondChanged = [bool](Invoke-Static $saveMethod @($workbook, $profile))
        if ($secondChanged) {
            throw "Save lan hai khong idempotent."
        }
        $workbook.Save()
    }

    if ($Reopen) {
        $workbook.Save()
        $workbook.Close($false)
        Release-ComObject $workbook
        $workbook = $null
        Start-Sleep -Milliseconds 500
        $workbook = $excel.Workbooks.Open($expectedPath, 0, $false)
    }

    $profile = Invoke-Static $loadMethod @($workbook)
    $payloadArguments = New-Object "object[]" 2
    $payloadArguments[0] = $workbook
    $payloadArguments[1] = $null
    if (-not [bool]$readPayloadMethod.Invoke($null, $payloadArguments)) {
        throw "Khong doc duoc project profile payload."
    }
    $payload = [string]$payloadArguments[1]
    $checksum = [string](Invoke-Static $checksumMethod @($payload))
    $roles = Invoke-Static $validateRolesMethod @($workbook)
    if (-not $roles.IsValid) {
        throw "Worksheet role regression."
    }

    [pscustomobject]@{
        Workbook = $workbook.FullName
        ProjectId = $profile.ProjectId
        SchemaVersion = $profile.SchemaVersion
        PreparedDate = $profile.PreparedDate
        ApprovalDate = $profile.ApprovalDate
        PriceDate = $profile.PriceDate
        RegulationPackageId = $profile.RegulationPackageId
        RegulationPackageVersion = $profile.RegulationPackageVersion
        RegulationPackageChecksum = $profile.RegulationPackageChecksum
        PriceProfileId = $profile.PriceProfileId
        OverrideLength = $profile.OverrideSummary.Length
        PayloadLength = $payload.Length
        PayloadChecksum = $checksum
        FirstSaveChanged = $firstChanged
        SecondSaveChanged = $secondChanged
        RolesValid = $roles.IsValid
        Saved = $workbook.Saved
    } | Format-List
}
finally {
    Release-OpenExcelWorkbookSession $openSession
}
