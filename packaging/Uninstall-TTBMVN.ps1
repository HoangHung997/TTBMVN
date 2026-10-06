param(
    [switch]$RemovePilotTrust,
    [switch]$Silent,
    [switch]$VerifyOnly
)

$ErrorActionPreference = "Stop"
$releaseInfoPath = Join-Path $PSScriptRoot "release-info.json"
$vstoInstaller = "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
$addinRegistryPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelAddIn1"
if (-not (Test-Path -LiteralPath $vstoInstaller)) { throw "Khong tim thay VSTOInstaller.exe." }

$installed = Get-ItemProperty -LiteralPath $addinRegistryPath -ErrorAction SilentlyContinue
if ($VerifyOnly) {
    [pscustomobject]@{
        Status = "PASS"
        Mode = "VERIFY-ONLY"
        Installed = ($null -ne $installed)
        Manifest = if ($null -eq $installed) { "" } else { [string]$installed.Manifest }
        ExcelOpen = [bool](Get-Process EXCEL -ErrorAction SilentlyContinue)
    } | Format-List
    return
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Hay dong tat ca cua so Excel truoc khi go add-in."
}
if ($null -ne $installed -and -not [string]::IsNullOrWhiteSpace([string]$installed.Manifest)) {
    $rawManifestValue = [string]$installed.Manifest
    $manifestValue = $rawManifestValue -replace '\|vstolocal$', ''
    $uri = $null
    $manifestPath = if ([Uri]::TryCreate($manifestValue, [UriKind]::Absolute, [ref]$uri) -and $uri.IsFile) {
        $uri.LocalPath
    }
    else { $manifestValue }
    $isDevelopmentRegistration =
        $rawManifestValue.EndsWith("|vstolocal", [StringComparison]::OrdinalIgnoreCase) -and
        $manifestPath -match '[\\/]bin[\\/]'
    if ($isDevelopmentRegistration -or -not (Test-Path -LiteralPath $manifestPath)) {
        Remove-Item -LiteralPath $addinRegistryPath -Recurse -Force
        Write-Host "Da go dang ky ExcelAddIn1 cu/stale."
    }
    else {
        $arguments = @("/Uninstall", ('"' + $manifestPath + '"'), "/Silent")
        $options = @{ FilePath=$vstoInstaller; ArgumentList=$arguments; Wait=$true; PassThru=$true; WindowStyle="Hidden" }
        $process = Start-Process @options
        if ($process.ExitCode -eq -401) {
            Remove-Item -LiteralPath $addinRegistryPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "VSTO khong con package; da don dang ky stale."
        }
        elseif ($process.ExitCode -ne 0) {
            throw "Go add-in that bai, exit=$($process.ExitCode)."
        }
        else {
            Write-Host "Da go TTBMVN Excel Tools."
        }
    }
}
else {
    Write-Host "TTBMVN Excel Tools khong dang duoc cai cho tai khoan nay."
}

if ($RemovePilotTrust) {
    if (-not (Test-Path -LiteralPath $releaseInfoPath)) { throw "Thieu release-info.json." }
    $releaseInfo = Get-Content -LiteralPath $releaseInfoPath -Raw | ConvertFrom-Json
    if ($releaseInfo.Channel -ne "PILOT") { throw "Chi go trust tu dong cho channel PILOT." }
    foreach ($store in @("Cert:\CurrentUser\Root", "Cert:\CurrentUser\TrustedPublisher")) {
        $certificate = Get-ChildItem -LiteralPath $store -ErrorAction SilentlyContinue |
            Where-Object Thumbprint -eq $releaseInfo.PublisherThumbprint |
            Select-Object -First 1
        if ($null -ne $certificate) { Remove-Item -LiteralPath $certificate.PSPath -Force }
    }
    Write-Host "Da go public pilot certificate khoi trust store cua tai khoan."
}
