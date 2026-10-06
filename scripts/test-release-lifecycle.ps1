param(
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$tmpRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "tmp"))
$acceptanceRoot = Join-Path $tmpRoot "DT-605-lifecycle"
$storeRoot = Join-Path $acceptanceRoot "package-store"
$fixture = Join-Path $repoRoot "Dutoanmau\Fixtures\RegulationPackages\BQP-RPBM-2025-2.0.0"
$update = Join-Path $repoRoot "tmp\DT-601-BQP-RPBM-2025-2.0.1.ttbupdate"
$evidencePath = Join-Path $repoRoot "tmp\DT-605-release-lifecycle.json"
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Assert-SafeTemporaryPath([string]$Path) {
    $resolved = [IO.Path]::GetFullPath($Path)
    $prefix = $tmpRoot.TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Duong dan lifecycle nam ngoai tmp: $resolved"
    }
}

function New-UpdateService([string]$Root) {
    $keys = [ExcelAddIn1.Core.OfflineUpdateTrustCatalog]::Production
    $keyEnumerableType = [System.Collections.Generic.IEnumerable[ExcelAddIn1.Core.OfflineUpdateTrustedKey]]
    $constructor = [ExcelAddIn1.Core.OfflineUpdateCenterService].GetConstructor(
        [type[]]@([string], $keyEnumerableType, [Version]))
    if ($null -eq $constructor) { throw "Khong tim thay OfflineUpdateCenterService constructor." }
    $arguments = New-Object "object[]" 3
    $arguments[0] = $Root
    $arguments[1] = $keys.PSObject.BaseObject
    $arguments[2] = ([Version]"1.0.0.0").PSObject.BaseObject
    return $constructor.Invoke($arguments)
}

Assert-SafeTemporaryPath $acceptanceRoot
if (Test-Path -LiteralPath $acceptanceRoot) {
    if (-not $Overwrite) {
        throw "Lifecycle root da ton tai: $acceptanceRoot. Dung -Overwrite."
    }
    Remove-Item -LiteralPath $acceptanceRoot -Recurse -Force
}
foreach ($required in @($fixture, $update, $corePath)) {
    if (-not (Test-Path -LiteralPath $required)) { throw "Thieu dau vao lifecycle: $required" }
}
New-Item -ItemType Directory -Path $acceptanceRoot -Force | Out-Null

[void][Reflection.Assembly]::LoadFrom($corePath)

# Clean install starts from an empty, isolated package store.
$store = [ExcelAddIn1.Core.RegulationPackageStore]::new($storeRoot)
Assert-True (@($store.ListInstalled()).Count -eq 0) "Clean store phai rong."
$oldInstall = $store.ImportFromDirectory($fixture)
Assert-True ($oldInstall.Status.ToString() -eq "Installed") "Khong clean-install duoc package 2.0.0."
Assert-True ($oldInstall.Package.DataVersion -eq "2.0.0") "Fixture upgrade khong phai 2.0.0."
$activation = [ExcelAddIn1.Core.RegulationPackageActivationStore]::new($storeRoot)
$activation.SetPreferred($oldInstall.Package, $store.ListInstalled())

# Upgrade uses the signed production update and must preserve the old installed version.
$service = New-UpdateService $storeRoot
$inspection = $service.Inspect($update, $oldInstall.Package)
Assert-True ($inspection.Verification.Disposition.ToString() -eq "Ready") "Update 2.0.1 khong o trang thai Ready."
Assert-True ($null -ne $inspection.Diff -and $inspection.Diff.HasChanges) "Upgrade khong co diff."
$upgrade = $service.InstallAndActivate($inspection)
$installedAfterUpgrade = @($service.ListInstalled())
$preferredAfterUpgrade = @($service.ListPreferred())
Assert-True ($upgrade.Status.ToString() -eq "InstalledAndActivated") "Upgrade khong duoc cai va kich hoat."
Assert-True ($installedAfterUpgrade.Count -eq 2) "Upgrade phai giu ca hai version da cai."
Assert-True ($preferredAfterUpgrade.Count -eq 1 -and $preferredAfterUpgrade[0].DataVersion -eq "2.0.1") `
    "Upgrade khong chon 2.0.1 lam preferred."

# A restart must reload the activation pointer from disk.
$restart = New-UpdateService $storeRoot
$preferredAfterRestart = @($restart.ListPreferred())
Assert-True ($preferredAfterRestart.Count -eq 1 -and $preferredAfterRestart[0].DataVersion -eq "2.0.1") `
    "Preferred package khong song qua restart."

# Uninstall is scoped to the isolated application-data store, then reinstall from signed media.
Assert-SafeTemporaryPath $storeRoot
Remove-Item -LiteralPath $storeRoot -Recurse -Force
$afterUninstall = New-UpdateService $storeRoot
Assert-True (@($afterUninstall.ListInstalled()).Count -eq 0) "Uninstall store khong sach."
$reinstallInspection = $afterUninstall.Inspect($update, $null)
$reinstall = $afterUninstall.InstallAndActivate($reinstallInspection)
$reinstalledPackages = @($afterUninstall.ListInstalled())
$reinstalledPreferred = @($afterUninstall.ListPreferred())
Assert-True ($reinstall.Status.ToString() -eq "InstalledAndActivated") "Reinstall khong thanh cong."
Assert-True ($reinstalledPackages.Count -eq 1 -and $reinstalledPackages[0].DataVersion -eq "2.0.1") `
    "Reinstall khong co dung package 2.0.1."
Assert-True ($reinstalledPreferred.Count -eq 1 -and $reinstalledPreferred[0].DataVersion -eq "2.0.1") `
    "Reinstall khong kich hoat package."

$evidence = [ordered]@{
    SchemaVersion = 1
    Task = "DT-605"
    Status = "PASS"
    StoreIsolation = $storeRoot
    CleanInstall = [ordered]@{
        Package = $oldInstall.Package.PackageId
        Version = $oldInstall.Package.DataVersion
        Checksum = $oldInstall.Package.PackageChecksum
        Status = $oldInstall.Status.ToString()
    }
    Upgrade = [ordered]@{
        TargetVersion = $upgrade.Package.DataVersion
        TargetChecksum = $upgrade.Package.PackageChecksum
        Disposition = $inspection.Verification.Disposition.ToString()
        DiffRows = $inspection.Diff.Changes.Count
        InstalledVersions = @($installedAfterUpgrade | ForEach-Object DataVersion)
        PreferredAfterRestart = $preferredAfterRestart[0].DataVersion
    }
    UninstallReinstall = [ordered]@{
        EmptyAfterUninstall = $true
        ReinstalledVersion = $reinstalledPackages[0].DataVersion
        PreferredVersion = $reinstalledPreferred[0].DataVersion
        Status = $reinstall.Status.ToString()
    }
}
[IO.File]::WriteAllText(
    $evidencePath,
    ($evidence | ConvertTo-Json -Depth 8),
    (New-Object Text.UTF8Encoding($false)))

$evidence | ConvertTo-Json -Depth 8
Write-Host "DT-605 release lifecycle PASS. Evidence: $evidencePath"
