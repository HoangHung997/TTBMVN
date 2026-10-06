param(
    [string]$Version = "1.0.0",
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "signing-certificate.ps1")
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
$solution = Join-Path $repoRoot "ExcelAddIn1.sln"
$publishSource = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\app.publish"
$artifactRoot = Join-Path $repoRoot "artifacts"
$certificate = Get-TtbmvnManifestSigningCertificate
$channel = Get-TtbmvnSigningChannel $certificate
if ($Version -notmatch '^[0-9]+\.[0-9]+\.[0-9]+$') { throw "Version phai dung dang major.minor.patch." }
$applicationVersion = $Version + ".0"
$releaseName = "TTBMVN-Excel-Tools-$Version-$channel-x64"
$releaseRoot = Join-Path $artifactRoot $releaseName
$archivePath = Join-Path $artifactRoot ($releaseName + ".zip")
$archiveHashPath = $archivePath + ".sha256"

function Assert-SafeChild([string]$Path, [string]$Parent) {
    $resolved = [IO.Path]::GetFullPath($Path)
    $resolvedParent = [IO.Path]::GetFullPath($Parent).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($resolvedParent, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Duong dan nam ngoai root cho phep: $resolved"
    }
}

Assert-SafeChild $releaseRoot $artifactRoot
Assert-SafeChild $publishSource (Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release")
foreach ($path in @($releaseRoot, $archivePath, $archiveHashPath)) {
    if (Test-Path -LiteralPath $path) {
        if (-not $Overwrite) { throw "Artifact da ton tai: $path. Dung -Overwrite." }
        if ((Get-Item -LiteralPath $path).PSIsContainer) { Remove-Item -LiteralPath $path -Recurse -Force }
        else { Remove-Item -LiteralPath $path -Force }
    }
}
if (Test-Path -LiteralPath $publishSource) { Remove-Item -LiteralPath $publishSource -Recurse -Force }
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null

& (Join-Path $PSScriptRoot "build-release.ps1") -Platform x64 -Configuration Release
if ($LASTEXITCODE -ne 0) { throw "Release build that bai." }
& $msbuild $solution /t:ExcelAddIn1:Publish `
    /p:Configuration=Release /p:Platform=x64 `
    /p:ApplicationVersion=$applicationVersion `
    /p:SignManifests=true `
    /p:ManifestCertificateThumbprint=$($certificate.Thumbprint) `
    /p:ManifestKeyFile= `
    /v:minimal
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $publishSource)) {
    throw "VSTO publish that bai."
}

New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null
Copy-Item -LiteralPath $publishSource -Destination (Join-Path $releaseRoot "installer") -Recurse
New-Item -ItemType Directory -Path (Join-Path $releaseRoot "certificate") -Force | Out-Null
$publicCertificatePath = Join-Path $releaseRoot "certificate\TTBMVN-Publisher.cer"
Export-Certificate -Cert $certificate -FilePath $publicCertificatePath -Type CERT | Out-Null

$setupPath = Join-Path $releaseRoot "installer\setup.exe"
$signatureOptions = @{ FilePath=$setupPath; Certificate=$certificate; HashAlgorithm="SHA256" }
if (-not [string]::IsNullOrWhiteSpace($env:TTBMVN_TIMESTAMP_URL)) {
    $signatureOptions.TimestampServer = $env:TTBMVN_TIMESTAMP_URL
}
$setupSignature = Set-AuthenticodeSignature @signatureOptions
if ($null -eq $setupSignature.SignerCertificate -or
    $setupSignature.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) {
    throw "Khong ky duoc setup.exe."
}

New-Item -ItemType Directory -Path (Join-Path $releaseRoot "updates"),(Join-Path $releaseRoot "sample"),(Join-Path $releaseRoot "docs") -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repoRoot "tmp\DT-601-BQP-RPBM-2025-2.0.1.ttbupdate") -Destination (Join-Path $releaseRoot "updates")
Copy-Item -LiteralPath (Join-Path $repoRoot "data\regulations\update-keys\TTBMVN-OFFLINE-2026-01.ttbpub") -Destination (Join-Path $releaseRoot "updates")
$sampleSource = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-605-end.xlsm"
$sampleName = "Du-toan-TP3-sample.xlsm"
Copy-Item -LiteralPath $sampleSource -Destination (Join-Path $releaseRoot "sample\$sampleName")
foreach ($file in @("Cai-dat-TTBMVN.cmd", "Install-TTBMVN.ps1", "Uninstall-TTBMVN.ps1", "Verify-TTBMVNRelease.ps1")) {
    Copy-Item -LiteralPath (Join-Path $repoRoot "packaging\$file") -Destination $releaseRoot
}
$releaseNotes = "RELEASE-NOTES-$Version.md"
if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "docs\du-toan\$releaseNotes"))) {
    throw "Thieu release notes cho version $Version."
}
foreach ($file in @("INSTALLATION.md", $releaseNotes, "RELEASE-GATE.md", "RELEASE-ACCEPTANCE.md", "COMPATIBILITY.md", "OFFLINE-UPDATE.md")) {
    Copy-Item -LiteralPath (Join-Path $repoRoot "docs\du-toan\$file") -Destination (Join-Path $releaseRoot "docs")
}

[void][Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"))
$packageBundle = [ExcelAddIn1.Core.RegulationPackageBundleReader]::Read(
    (Join-Path $repoRoot "data\regulations\packages\BQP-RPBM-2025\bundle"))
$packageChecksum = $packageBundle.Package.PackageChecksum
$info = [ordered]@{
    SchemaVersion = 1
    Product = "TTBMVN Excel Tools"
    Version = $Version
    ApplicationVersion = $applicationVersion
    Channel = $channel
    Platform = "x64"
    MinimumDotNet = "4.8.1"
    PublisherSubject = $certificate.Subject
    PublisherThumbprint = $certificate.Thumbprint
    PublisherNotAfter = $certificate.NotAfter.ToUniversalTime().ToString("o")
    Timestamped = -not [string]::IsNullOrWhiteSpace($env:TTBMVN_TIMESTAMP_URL)
    SetupSignatureStatusAtBuild = $setupSignature.Status.ToString()
    AcceptanceStatus = "DT-605 14/14 PASS; DT-701 PASS"
    AcceptanceCheckpointSha256 = "56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C"
    RegulationPackage = "BQP-RPBM-2025@2.0.1"
    RegulationPackageChecksum = $packageChecksum
    SampleWorkbookPath = "sample/$sampleName"
    SampleWorkbookSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $sampleSource).Hash
    ReleaseArchiveSha256 = "SEE-SIDECAR"
    CreatedAtUtc = [DateTime]::UtcNow.ToString("o")
}
[IO.File]::WriteAllText(
    (Join-Path $releaseRoot "release-info.json"),
    ($info | ConvertTo-Json -Depth 6),
    (New-Object Text.UTF8Encoding($false)))

$checksumLines = Get-ChildItem -LiteralPath $releaseRoot -Recurse -File |
    Where-Object Name -ne "SHA256SUMS.txt" |
    ForEach-Object {
        $relative = $_.FullName.Substring($releaseRoot.Length).TrimStart('\').Replace('\', '/')
        (Get-FileHash -Algorithm SHA256 -LiteralPath $_.FullName).Hash + "`t" + $relative
    } | Sort-Object
[IO.File]::WriteAllLines(
    (Join-Path $releaseRoot "SHA256SUMS.txt"),
    $checksumLines,
    (New-Object Text.UTF8Encoding($false)))

& (Join-Path $releaseRoot "Verify-TTBMVNRelease.ps1") `
    -ReleaseRoot $releaseRoot `
    -EvidencePath (Join-Path $repoRoot "tmp\DT-606-release-verification.json")
if ($LASTEXITCODE -ne 0) { throw "Release verification that bai." }

Compress-Archive -Path (Join-Path $releaseRoot "*") -DestinationPath $archivePath -CompressionLevel Optimal
$archiveHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $archivePath).Hash
[IO.File]::WriteAllText($archiveHashPath, $archiveHash + "  " + [IO.Path]::GetFileName($archivePath) + "`n", (New-Object Text.UTF8Encoding($false)))
[pscustomobject]@{
    ReleaseDirectory = $releaseRoot
    Archive = $archivePath
    ArchiveSha256 = $archiveHash
    Channel = $channel
    Publisher = $certificate.Subject
    Thumbprint = $certificate.Thumbprint
    SetupSignature = $setupSignature.Status.ToString()
} | Format-List
