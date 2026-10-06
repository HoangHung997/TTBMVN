param(
    [string]$Version = "1.0.0",
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$channel = "PILOT"
$releaseName = "TTBMVN-Excel-Tools-$Version-$channel-x64"
$archive = Join-Path $repoRoot "artifacts\$releaseName.zip"
$sidecar = $archive + ".sha256"
$tmpRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "tmp"))
$extractRoot = Join-Path $tmpRoot "DT-606-gate-extracted"
$tamperRoot = Join-Path $tmpRoot "DT-606-gate-tampered"
$evidencePath = Join-Path $tmpRoot "DT-606-release-gate.json"
$acceptancePath = Join-Path $tmpRoot "DT-605-release-acceptance.json"

function Assert-SafeTemporaryPath([string]$Path) {
    $resolved = [IO.Path]::GetFullPath($Path)
    if (-not $resolved.StartsWith($tmpRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Duong dan gate nam ngoai tmp: $resolved"
    }
}

foreach ($required in @($archive, $sidecar, $acceptancePath)) {
    if (-not (Test-Path -LiteralPath $required)) { throw "Gate thieu dau vao: $required" }
}
foreach ($path in @($extractRoot, $tamperRoot)) {
    Assert-SafeTemporaryPath $path
    if (Test-Path -LiteralPath $path) {
        if (-not $Overwrite) { throw "Thu muc gate da ton tai: $path. Dung -Overwrite." }
        Remove-Item -LiteralPath $path -Recurse -Force
    }
}

& (Join-Path $PSScriptRoot "build-release.ps1") -Platform x64 -Configuration Release
if ($LASTEXITCODE -ne 0) { throw "Release build gate that bai." }

$expectedArchiveHash = ((Get-Content -LiteralPath $sidecar -Raw).Trim() -split '\s+')[0]
$actualArchiveHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $archive).Hash
if ($actualArchiveHash -ne $expectedArchiveHash) { throw "Archive SHA-256 khong khop sidecar." }

Expand-Archive -LiteralPath $archive -DestinationPath $extractRoot
$verifier = Join-Path $extractRoot "Verify-TTBMVNRelease.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $verifier `
    -ReleaseRoot $extractRoot `
    -EvidencePath (Join-Path $tmpRoot "DT-606-release-zip-verification.json")
if ($LASTEXITCODE -ne 0) { throw "Verifier release ZIP that bai." }
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $extractRoot "Install-TTBMVN.ps1") -VerifyOnly
if ($LASTEXITCODE -ne 0) { throw "Install VerifyOnly that bai." }
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $extractRoot "Uninstall-TTBMVN.ps1") -VerifyOnly
if ($LASTEXITCODE -ne 0) { throw "Uninstall VerifyOnly that bai." }

New-Item -ItemType Directory -Path $tamperRoot -Force | Out-Null
Copy-Item -Path (Join-Path $extractRoot "*") -Destination $tamperRoot -Recurse
$tamperTarget = Join-Path $tamperRoot "docs\RELEASE-NOTES-$Version.md"
[IO.File]::AppendAllText($tamperTarget, "`nTAMPER-DT-606`n", (New-Object Text.UTF8Encoding($false)))
$tamperOutput = Join-Path $tmpRoot "DT-606-tamper-verifier.out.txt"
$tamperError = Join-Path $tmpRoot "DT-606-tamper-verifier.err.txt"
$tamperArguments = ('-NoProfile -ExecutionPolicy Bypass -File "' +
    (Join-Path $tamperRoot "Verify-TTBMVNRelease.ps1") + '" -ReleaseRoot "' + $tamperRoot + '"')
$tamperProcess = Start-Process -FilePath powershell.exe `
    -ArgumentList $tamperArguments `
    -Wait -PassThru -WindowStyle Hidden `
    -RedirectStandardOutput $tamperOutput `
    -RedirectStandardError $tamperError
$tamperMessage = (Get-Content -LiteralPath $tamperOutput,$tamperError -Raw) -join "`n"
if ($tamperProcess.ExitCode -eq 0 -or -not $tamperMessage.Contains("Checksum sai")) {
    throw "Verifier khong fail-closed khi release bi sua."
}

$acceptance = Get-Content -LiteralPath $acceptancePath -Raw | ConvertFrom-Json
if ($acceptance.Status -ne "PASS" -or $acceptance.ExecutedTestCount -ne 14 -or
    $acceptance.CheckpointSha256 -ne "56C3437B63E84ECEF2C1CEACEAE72421285571CCF083C99A626047D9C1DBF53C") {
    throw "DT-605 acceptance evidence khong dat release gate."
}
if (Get-ChildItem -LiteralPath $repoRoot -Recurse -File -Filter "*.pfx") {
    throw "Workspace con chua PFX/private signing material."
}

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "test-workbook-checkpoint.ps1") `
    -CheckpointPath ".\Dutoanmau\Checkpoints\DT-606-start.xlsm" `
    -ExpectedE27 1201557000 `
    -ExpectedFormulaErrors 92
if ($LASTEXITCODE -ne 0) { throw "DT-606 semantic checkpoint that bai." }

$releaseInfo = Get-Content -LiteralPath (Join-Path $extractRoot "release-info.json") -Raw | ConvertFrom-Json
$result = [ordered]@{
    SchemaVersion = 1
    Task = "DT-606"
    Status = "PASS"
    Version = $releaseInfo.Version
    Channel = $releaseInfo.Channel
    Platform = $releaseInfo.Platform
    Archive = $archive
    ArchiveSha256 = $actualArchiveHash
    PublisherSubject = $releaseInfo.PublisherSubject
    PublisherThumbprint = $releaseInfo.PublisherThumbprint
    ReleaseFiles = (Get-ChildItem -LiteralPath $extractRoot -Recurse -File).Count
    ManifestVerification = "PASS"
    InstallVerifyOnly = "PASS"
    UninstallVerifyOnly = "PASS"
    TamperRejected = "PASS"
    PrivatePfxInWorkspace = "ABSENT"
    Acceptance = "14/14 PASS"
    SemanticCheckpoint = "PASS"
    ActualVstoMutation = "NOT_RUN_USER_EXCEL_OPEN"
}
[IO.File]::WriteAllText(
    $evidencePath,
    ($result | ConvertTo-Json -Depth 6),
    (New-Object Text.UTF8Encoding($false)))
$result | ConvertTo-Json -Depth 6
Write-Host "DT-606 RELEASE GATE PASSED."
