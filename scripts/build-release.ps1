param(
    [string]$Platform = "x64",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $repoRoot "ExcelAddIn1.sln"
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = $null
. (Join-Path $PSScriptRoot "signing-certificate.ps1")

if (Test-Path $vswhere) {
    $msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild `
        -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
}

if ([string]::IsNullOrWhiteSpace($msbuild) -or -not (Test-Path $msbuild)) {
    $command = Get-Command msbuild.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        $msbuild = $command.Source
    }
}

if ([string]::IsNullOrWhiteSpace($msbuild) -or -not (Test-Path $msbuild)) {
    throw "MSBuild not found. Install Visual Studio with the Office/VSTO workload."
}

$signingCertificate = Get-TtbmvnManifestSigningCertificate
& $msbuild $solution /t:Build /p:Configuration=$Configuration /p:Platform=$Platform `
    /p:SignManifests=true `
    /p:ManifestCertificateThumbprint=$($signingCertificate.Thumbprint) `
    /p:ManifestKeyFile= `
    /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
}

$testExe = Join-Path $repoRoot "ExcelAddIn1.Tests\bin\$Configuration\ExcelAddIn1.Tests.exe"
if (Test-Path $testExe) {
    & $testExe
    if ($LASTEXITCODE -ne 0) {
        throw "Tests failed."
    }
}

Write-Host ("Release build completed: $Configuration|$Platform; Signing=" +
    (Get-TtbmvnSigningChannel $signingCertificate) + "; Thumbprint=" +
    $signingCertificate.Thumbprint)
