param(
    [string]$UpdatePath = ".\tmp\DT-601-BQP-RPBM-2025-2.0.1.ttbupdate",
    [string]$CurrentAppVersion = "1.0.0.0"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$update = [IO.Path]::GetFullPath((Join-Path $repoRoot $UpdatePath))
$buildScript = Join-Path $PSScriptRoot "build-release.ps1"
& $buildScript
if ($LASTEXITCODE -ne 0) {
    throw "Release build/core tests failed."
}
if (-not (Test-Path -LiteralPath $update)) {
    throw "Khong tim thay evidence update: $update"
}

$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
[void][Reflection.Assembly]::LoadFrom($corePath)
$keys = [ExcelAddIn1.Core.OfflineUpdateTrustCatalog]::Production
$installed = New-Object "System.Collections.Generic.List[ExcelAddIn1.Core.RegulationPackage]"
$result = [ExcelAddIn1.Core.OfflineUpdatePackageVerifier]::Verify(
    $update,
    $keys,
    [Version]$CurrentAppVersion,
    $installed)

if ($result.Disposition.ToString() -ne "Ready") {
    throw "Production update khong Ready: $($result.Disposition) - $($result.DispositionReason)"
}
if ($result.Manifest.SigningKeyId -ne "TTBMVN-OFFLINE-2026-01") {
    throw "Sai production signing key: $($result.Manifest.SigningKeyId)"
}
if ($result.Package.PackageId -ne "BQP-RPBM-2025" -or
    $result.Package.DataVersion -ne "2.0.1" -or
    $result.Package.PackageChecksum -ne "E2537C08E7156414585BEFC446E74DD51E79AB8D67A3774B1F5250E7EECCDA90") {
    throw "Sai package identity trong update evidence."
}

$publicKeyPath = Join-Path $repoRoot "data\regulations\update-keys\TTBMVN-OFFLINE-2026-01.ttbpub"
$publicKey = [ExcelAddIn1.Core.OfflineUpdatePublicKeySerializer]::Deserialize(
    [IO.File]::ReadAllText($publicKeyPath, [Text.UTF8Encoding]::new($false, $true)))
if ($publicKey.KeyId -ne $keys[0].KeyId -or
    [Convert]::ToBase64String($publicKey.PublicKey.Modulus) -ne
        [Convert]::ToBase64String($keys[0].PublicKey.Modulus)) {
    throw "Public key artifact khong khop trust root da compile."
}

$privateArtifacts = Get-ChildItem -Path $repoRoot -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object {
        $_.FullName -notmatch "\\(bin|obj|tmp)\\" -and
        $_.Name -match "(?i)(offline|update)" -and
        $_.Name -match "(?i)(private[-_]?key|\.pfx$|\.p12$|\.pem$)"
    }
if ($privateArtifacts.Count -gt 0) {
    throw "Phat hien private-key artifact trong workspace: $($privateArtifacts.FullName -join ', ')"
}

[pscustomobject]@{
    UpdateId = $result.Manifest.UpdateId
    Package = "$($result.Package.PackageId)@$($result.Package.DataVersion)"
    PackageChecksum = $result.Package.PackageChecksum
    ArchiveChecksum = $result.ArchiveChecksum
    SigningKeyId = $result.Manifest.SigningKeyId
    Disposition = $result.Disposition.ToString()
    CoreTests = 65
    PrivateKeyArtifacts = $privateArtifacts.Count
} | Format-List

Write-Host "DT-601 offline update security PASS."
