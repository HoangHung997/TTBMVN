param(
    [string]$ReleaseRoot = $PSScriptRoot,
    [string]$EvidencePath
)

$ErrorActionPreference = "Stop"
$root = [IO.Path]::GetFullPath($ReleaseRoot)
$infoPath = Join-Path $root "release-info.json"
$checksumPath = Join-Path $root "SHA256SUMS.txt"
$certificatePath = Join-Path $root "certificate\TTBMVN-Publisher.cer"
$setupPath = Join-Path $root "installer\setup.exe"
$vstoPath = Join-Path $root "installer\ExcelAddIn1.vsto"
$mageCandidates = @(
    "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\mage.exe",
    "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe"
)
$mage = $mageCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
foreach ($required in @($infoPath, $checksumPath, $certificatePath, $setupPath, $vstoPath, $mage)) {
    if ([string]::IsNullOrWhiteSpace([string]$required) -or -not (Test-Path -LiteralPath $required)) {
        throw "Release thieu file bat buoc: $required"
    }
}

$info = Get-Content -LiteralPath $infoPath -Raw | ConvertFrom-Json
$expected = @{}
foreach ($line in Get-Content -LiteralPath $checksumPath) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $parts = $line -split "`t", 2
    if ($parts.Count -ne 2 -or $expected.ContainsKey($parts[1])) { throw "SHA256SUMS.txt khong hop le." }
    $expected[$parts[1]] = $parts[0]
}
$actualFiles = @(Get-ChildItem -LiteralPath $root -Recurse -File |
    Where-Object { $_.FullName -ne $checksumPath })
foreach ($file in $actualFiles) {
    $relative = $file.FullName.Substring($root.Length).TrimStart('\').Replace('\', '/')
    if (-not $expected.ContainsKey($relative)) { throw "File release khong co checksum: $relative" }
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $file.FullName).Hash
    if ($hash -ne $expected[$relative]) { throw "Checksum sai: $relative" }
    [void]$expected.Remove($relative)
}
if ($expected.Count -ne 0) { throw "SHA256SUMS.txt co file thieu trong release." }

$certificate = New-Object Security.Cryptography.X509Certificates.X509Certificate2($certificatePath)
if ($certificate.Thumbprint -ne $info.PublisherThumbprint) { throw "Publisher certificate/thumbprint khong khop." }
$setupSignature = Get-AuthenticodeSignature -FilePath $setupPath
if ($null -eq $setupSignature.SignerCertificate -or
    $setupSignature.SignerCertificate.Thumbprint -ne $info.PublisherThumbprint) {
    throw "setup.exe khong duoc ky boi publisher da khai bao."
}
if ($info.Channel -ne "PILOT" -and $setupSignature.Status.ToString() -ne "Valid") {
    throw "setup.exe channel ngoai PILOT phai co trust status Valid."
}

& $mage -Verify $vstoPath | Out-Host
if ($LASTEXITCODE -ne 0) { throw "Deployment manifest signature khong hop le." }
$applicationManifests = @(Get-ChildItem -LiteralPath (Join-Path $root "installer\Application Files") `
    -Recurse -Filter "ExcelAddIn1.dll.manifest")
if ($applicationManifests.Count -ne 1) { throw "Release phai co dung mot application manifest." }
$applicationManifest = $applicationManifests[0]
& $mage -Verify $applicationManifest.FullName | Out-Host
if ($LASTEXITCODE -ne 0) { throw "Application manifest signature khong hop le." }
$vstoText = Get-Content -LiteralPath $vstoPath -Raw
if (-not $vstoText.Contains($info.PublisherSubject) -or -not $vstoText.Contains('processorArchitecture="amd64"')) {
    throw "Deployment manifest sai publisher hoac architecture."
}
if (Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object Extension -eq ".pfx") {
    throw "Release khong duoc chua private PFX."
}

$samplePath = Join-Path $root $info.SampleWorkbookPath
if ((Get-FileHash -Algorithm SHA256 -LiteralPath $samplePath).Hash -ne $info.SampleWorkbookSha256) {
    throw "Workbook mau sai checksum."
}
$result = [ordered]@{
    Status = "PASS"
    Product = $info.Product
    Version = $info.Version
    Channel = $info.Channel
    Platform = $info.Platform
    PublisherSubject = $info.PublisherSubject
    PublisherThumbprint = $info.PublisherThumbprint
    SetupSignatureStatus = $setupSignature.Status.ToString()
    FilesVerified = $actualFiles.Count
    DeploymentManifest = "PASS"
    ApplicationManifest = "PASS"
    SampleWorkbook = "PASS"
}
if (-not [string]::IsNullOrWhiteSpace($EvidencePath)) {
    [IO.File]::WriteAllText(
        [IO.Path]::GetFullPath($EvidencePath),
        ($result | ConvertTo-Json -Depth 5),
        (New-Object Text.UTF8Encoding($false)))
}
$result | ConvertTo-Json -Depth 5
