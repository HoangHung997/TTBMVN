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
$expectedPath = [IO.Path]::GetFullPath($WorkbookPath)
$integrationRoot = [IO.Path]::GetFullPath(
    (Join-Path $repoRoot "Dutoanmau\DT204Integration"))
$variantRoot = [IO.Path]::GetFullPath(
    (Join-Path $repoRoot "Dutoanmau\Variants"))
$workingCopyPath = Join-Path $variantRoot "DT-204-working-copy.xlsm"
$machineCopyPath = Join-Path $variantRoot "DT-204-machine-copy.xlsm"
$emptyStore = Join-Path $integrationRoot "empty-store"
$defaultStore = Join-Path $env:LOCALAPPDATA "TTBMVNExcelTools\RegulationPackages"
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"

function Assert-Descendant([string]$Path, [string]$Parent) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    $fullParent = [IO.Path]::GetFullPath($Parent).TrimEnd('\') + '\'
    if (-not $fullPath.StartsWith($fullParent, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path nam ngoai thu muc test: $fullPath"
    }
}

Assert-Descendant $integrationRoot (Join-Path $repoRoot "Dutoanmau")
Assert-Descendant $workingCopyPath $variantRoot
Assert-Descendant $machineCopyPath $variantRoot
if (Test-Path $integrationRoot) {
    if (-not $Overwrite) {
        throw "Thu muc test DT-204 da ton tai. Dung -Overwrite de tao lai."
    }
    Remove-Item -LiteralPath $integrationRoot -Recurse -Force
}
foreach ($path in @($workingCopyPath, $machineCopyPath)) {
    if (Test-Path $path) {
        if (-not $Overwrite) {
            throw "Variant DT-204 da ton tai. Dung -Overwrite de tao lai."
        }
        Remove-Item -LiteralPath $path -Force
    }
}
New-Item -ItemType Directory -Path $integrationRoot,$variantRoot,$emptyStore -Force | Out-Null
Copy-Item -LiteralPath $expectedPath -Destination $workingCopyPath

[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
if (-not ("TTBMVN.Tests.PackagePinIntegrationSupport" -as [type])) {
    Add-Type -ReferencedAssemblies @($corePath, "System.dll", "System.Core.dll") -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelAddIn1.Core;

namespace TTBMVN.Tests
{
    public static class PackagePinIntegrationSupport
    {
        public static RegulationPackageInstallResult CreateAndInstall(
            string bundleRoot,
            string storeRoot,
            string packageId,
            string dataVersion,
            string marker)
        {
            string bundle = Path.Combine(bundleRoot, "bundle-" + dataVersion);
            string modulesDirectory = Path.Combine(bundle, RegulationPackageLayout.ModulesDirectoryName);
            Directory.CreateDirectory(modulesDirectory);
            var modules = new List<RegulationPackageModuleManifest>();
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                string content = "DT204|" + dataVersion + "|" + kind + "|" + marker;
                File.WriteAllText(
                    Path.Combine(modulesDirectory, RegulationPackageLayout.GetModuleFileName(kind)),
                    content,
                    new UTF8Encoding(false));
                modules.Add(new RegulationPackageModuleManifest(
                    kind,
                    "DT204-" + kind.ToString().ToUpperInvariant(),
                    1,
                    dataVersion,
                    content.Length,
                    RegulationPackageSerializer.ComputeSha256(content)));
            }

            var sources = new[]
            {
                new RegulationPackageSourceDocument(
                    "DT204-TEST-SOURCE",
                    "DT-204 persistence test source - not for calculation",
                    "TTBMVN Tests",
                    new DateTime(2026, 8, 3),
                    new DateTime(2026, 8, 3),
                    null,
                    "https://example.invalid/ttbmvn/dt204",
                    new string('A', 64))
            };
            RegulationPackage package = RegulationPackage.Create(
                packageId,
                dataVersion,
                new DateTime(2026, 8, 3),
                null,
                RegulationPackageStatus.Published,
                "DT-204 persistence test package; not for calculation.",
                sources,
                modules);
            File.WriteAllText(
                Path.Combine(bundle, RegulationPackageLayout.ManifestFileName),
                RegulationPackageSerializer.Serialize(package),
                new UTF8Encoding(false));
            return new RegulationPackageStore(storeRoot).ImportFromDirectory(bundle);
        }
    }
}
"@
}

$profileServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$pinServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookRegulationPackageService", $true)
$roleServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$loadProfileMethod = $profileServiceType.GetMethod("LoadRequired")
$readPayloadMethod = $profileServiceType.GetMethod("TryReadPayload")
$pinMethod = $pinServiceType.GetMethod("PinAndSave")
$verifyMethod = $pinServiceType.GetMethod("Verify")
$validateRolesMethod = $roleServiceType.GetMethod("Validate")

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

function Find-Workbook([object]$Excel, [string]$FullPath) {
    $books = $Excel.Workbooks
    try {
        for ($index = 1; $index -le $books.Count; $index++) {
            $candidate = $books.Item($index)
            if ([string]::Equals($candidate.FullName, $FullPath, [StringComparison]::OrdinalIgnoreCase)) {
                return $candidate
            }
            Release-ComObject $candidate
        }
    }
    finally {
        Release-ComObject $books
    }
    return $null
}

$packageId = "TEST-DT103-PACKAGE-NOT-FOR-CALCULATION"
$version1 = [TTBMVN.Tests.PackagePinIntegrationSupport]::CreateAndInstall(
    $integrationRoot, $defaultStore, $packageId, "1.0.0", "PINNED")
$version2 = [TTBMVN.Tests.PackagePinIntegrationSupport]::CreateAndInstall(
    $integrationRoot, $defaultStore, $packageId, "2.0.0", "NEWER")

$excelSession = Start-IsolatedExcelTestProcess
$excel = $excelSession.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $excel.Workbooks.Open($workingCopyPath, 0, $false)
$verifyBook = $null
$copy = $null
try {
    $pinArgs = New-Object "object[]" 2
    $pinArgs[0] = $workbook.PSObject.BaseObject
    $pinArgs[1] = $version1.Package.PSObject.BaseObject
    [void]$pinMethod.Invoke($null, $pinArgs)
    $workbook.Save()
    $workbook.SaveCopyAs($machineCopyPath)

    $payloadArgs = New-Object "object[]" 2
    $payloadArgs[0] = $workbook.PSObject.BaseObject
    $payloadArgs[1] = $null
    if (-not [bool]$readPayloadMethod.Invoke($null, $payloadArgs)) {
        throw "Khong doc duoc profile payload sau pin."
    }

    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null
    $verifyBook = $excel.Workbooks.Open($workingCopyPath, 0, $true)

    $profile = Invoke-Static $loadProfileMethod @($verifyBook)
    if ($profile.SchemaVersion -ne 2 -or
        $profile.RegulationPackageVersion -ne "1.0.0" -or
        $profile.RegulationPackageChecksum -ne $version1.Package.PackageChecksum) {
        throw "ProjectProfile pin khong ton tai sau save/reopen."
    }
    $verified = Invoke-Static $verifyMethod @($verifyBook, $defaultStore)
    if ($verified.Status.ToString() -ne "Available" -or
        $verified.Package.DataVersion -ne "1.0.0") {
        throw "Default store khong nap dung package v1 da pin."
    }
    if ($version2.Package.DataVersion -ne "2.0.0") {
        throw "Khong cai duoc package newer de test pin."
    }

    $roles = Invoke-Static $validateRolesMethod @($verifyBook)
    if (-not $roles.IsValid) {
        throw "Worksheet role regression."
    }
    $summary = $verifyBook.Worksheets.Item("THKP-TC")
    try {
        if ([double]$summary.Range("E27").Value2 -ne 1198731000) {
            throw "Baseline E27 regression."
        }
    }
    finally {
        Release-ComObject $summary
    }

    $verifyBook.Close($false)
    Release-ComObject $verifyBook
    $verifyBook = $null

    $copy = $excel.Workbooks.Open($machineCopyPath, 0, $true)
    $copyProfile = Invoke-Static $loadProfileMethod @($copy)
    $missing = Invoke-Static $verifyMethod @($copy, $emptyStore)
    if ($missing.Status.ToString() -ne "Missing" -or
        $copyProfile.RegulationPackageChecksum -ne $version1.Package.PackageChecksum) {
        throw "Machine-copy khong bao Missing dung package da pin."
    }
    $copyAvailable = Invoke-Static $verifyMethod @($copy, $defaultStore)
    if ($copyAvailable.Status.ToString() -ne "Available" -or
        $copyAvailable.Package.DataVersion -ne "1.0.0") {
        throw "Machine-copy khong nap duoc package sau khi co store."
    }

    [pscustomobject]@{
        Workbook = $workingCopyPath
        Schema = $profile.SchemaVersion
        PackageId = $profile.RegulationPackageId
        PinnedVersion = $profile.RegulationPackageVersion
        PinnedChecksum = $profile.RegulationPackageChecksum
        NewerInstalled = $version2.Package.DataVersion
        DefaultStore = $verified.Status
        EmptyStore = $missing.Status
        MachineCopy = $machineCopyPath
        ProfilePayloadLength = ([string]$payloadArgs[1]).Length
        Roles = "PASS"
        E27 = 1198731000
    } | Format-List
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    if ($null -ne $verifyBook) {
        $verifyBook.Close($false)
        Release-ComObject $verifyBook
    }
    if ($null -ne $copy) {
        $copy.Close($false)
        Release-ComObject $copy
    }
    Stop-IsolatedExcelTestProcess $excelSession
}

Get-FileHash -LiteralPath $machineCopyPath -Algorithm SHA256 | Format-List
