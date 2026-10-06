param(
    [switch]$TrustPilotCertificate,
    [switch]$Silent,
    [switch]$VerifyOnly
)

$ErrorActionPreference = "Stop"
$releaseRoot = $PSScriptRoot
$releaseInfoPath = Join-Path $releaseRoot "release-info.json"
$vstoPath = Join-Path $releaseRoot "installer\ExcelAddIn1.vsto"
$certificatePath = Join-Path $releaseRoot "certificate\TTBMVN-Publisher.cer"
$vstoInstaller = "C:\Program Files\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
$addinRegistryPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\ExcelAddIn1"
$staleRegistrationBackup = $null

function Invoke-VstoInstaller(
    [string]$Operation,
    [string]$ManifestPath,
    [switch]$AllowNotInstalled,
    [switch]$ForceSilent) {
    $arguments = @($Operation, ('"' + $ManifestPath + '"'))
    if ($Silent -or $ForceSilent) { $arguments += "/Silent" }
    $options = @{
        FilePath = $vstoInstaller
        ArgumentList = $arguments
        Wait = $true
        PassThru = $true
    }
    if ($Silent -or $ForceSilent) { $options.WindowStyle = "Hidden" }
    $process = Start-Process @options
    if ($process.ExitCode -ne 0) {
        if ($AllowNotInstalled -and $process.ExitCode -eq -401) {
            return $false
        }
        throw "VSTOInstaller $Operation that bai, exit=$($process.ExitCode)."
    }
    return $true
}

function Remove-StaleAddinRegistration([object]$Registration, [string]$Reason) {
    $backupRoot = Join-Path $env:LOCALAPPDATA "TTBMVNExcelTools\InstallBackups"
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
    $backup = [ordered]@{
        CapturedAtUtc = [DateTime]::UtcNow.ToString("o")
        Reason = $Reason
        RegistryPath = $addinRegistryPath
        FriendlyName = [string]$Registration.FriendlyName
        Description = [string]$Registration.Description
        Manifest = [string]$Registration.Manifest
        LoadBehavior = [int]$Registration.LoadBehavior
    }
    $backupPath = Join-Path $backupRoot (
        "stale-registration-" + [DateTime]::UtcNow.ToString("yyyyMMdd-HHmmss-fff") + ".json")
    [IO.File]::WriteAllText(
        $backupPath,
        ($backup | ConvertTo-Json -Depth 4),
        (New-Object Text.UTF8Encoding($false)))
    $script:staleRegistrationBackup = $backup
    Remove-Item -LiteralPath $addinRegistryPath -Recurse -Force
    Write-Host "Da sao luu va go dang ky add-in cu: $Reason"
}

function Restore-StaleAddinRegistration {
    if ($null -eq $script:staleRegistrationBackup -or (Test-Path -LiteralPath $addinRegistryPath)) {
        return
    }
    New-Item -Path $addinRegistryPath -Force | Out-Null
    New-ItemProperty -LiteralPath $addinRegistryPath -Name "FriendlyName" `
        -Value ([string]$script:staleRegistrationBackup.FriendlyName) -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $addinRegistryPath -Name "Description" `
        -Value ([string]$script:staleRegistrationBackup.Description) -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $addinRegistryPath -Name "Manifest" `
        -Value ([string]$script:staleRegistrationBackup.Manifest) -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $addinRegistryPath -Name "LoadBehavior" `
        -Value ([int]$script:staleRegistrationBackup.LoadBehavior) -PropertyType DWord -Force | Out-Null
    Write-Host "Da khoi phuc dang ky add-in cu do cai ban moi that bai."
}

function Convert-ManifestValueToPath([string]$Value) {
    $clean = if ($null -eq $Value) { "" } else { $Value }
    $clean = $clean -replace '\|vstolocal$', ''
    if ([string]::IsNullOrWhiteSpace($clean)) { return "" }
    $uri = $null
    if ([Uri]::TryCreate($clean, [UriKind]::Absolute, [ref]$uri) -and $uri.IsFile) {
        return $uri.LocalPath
    }
    return $clean
}

function Get-InstalledProductSubscriptions([string]$ProductName) {
    $uninstallRoot = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    if (-not (Test-Path -LiteralPath $uninstallRoot)) { return }
    foreach ($key in Get-ChildItem -LiteralPath $uninstallRoot -ErrorAction SilentlyContinue) {
        $item = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
        if ($null -eq $item -or
            -not [string]::Equals([string]$item.DisplayName, $ProductName, [StringComparison]::OrdinalIgnoreCase) -or
            -not [string]::Equals([string]$item.Publisher, "TTBMVN", [StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        [pscustomobject]@{
            RegistryPath = $key.PSPath
            DisplayName = [string]$item.DisplayName
            DisplayVersion = [string]$item.DisplayVersion
            Manifest = [string]$item.UrlUpdateInfo
            UninstallString = [string]$item.UninstallString
        }
    }
}

function Remove-InstalledProductSubscriptions([string]$ProductName) {
    foreach ($subscription in @(Get-InstalledProductSubscriptions $ProductName)) {
        $manifestValue = $subscription.Manifest
        if ([string]::IsNullOrWhiteSpace($manifestValue) -and
            $subscription.UninstallString -match '(?i)/Uninstall\s+("[^"]+\.vsto"|\S+\.vsto)') {
            $manifestValue = $Matches[1].Trim('"')
        }
        if ([string]::IsNullOrWhiteSpace($manifestValue)) {
            throw "Khong doc duoc manifest go cai cua $($subscription.DisplayName) $($subscription.DisplayVersion)."
        }
        $manifestPath = Convert-ManifestValueToPath $manifestValue
        $uninstallTarget = if (Test-Path -LiteralPath $manifestPath) { $manifestPath } else { $manifestValue }
        Write-Host "Dang go subscription $($subscription.DisplayName) $($subscription.DisplayVersion)..."
        $removed = Invoke-VstoInstaller "/Uninstall" $uninstallTarget -AllowNotInstalled -ForceSilent
        if (-not $removed) {
            Remove-Item -LiteralPath $subscription.RegistryPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Subscription registry da stale; da don entry cu."
        }
    }
}

foreach ($required in @($releaseInfoPath, $vstoPath, $vstoInstaller)) {
    if (-not (Test-Path -LiteralPath $required)) { throw "Thieu file cai dat: $required" }
}
$releaseInfo = Get-Content -LiteralPath $releaseInfoPath -Raw | ConvertFrom-Json
$existing = Get-ItemProperty -LiteralPath $addinRegistryPath -ErrorAction SilentlyContinue
if ($VerifyOnly) {
    if ($releaseInfo.Channel -eq "PILOT") {
        if (-not (Test-Path -LiteralPath $certificatePath)) { throw "Thieu public pilot certificate." }
        $certificate = New-Object Security.Cryptography.X509Certificates.X509Certificate2($certificatePath)
        if ($certificate.Thumbprint -ne $releaseInfo.PublisherThumbprint) {
            throw "Public certificate khong khop release-info.json."
        }
    }
    [pscustomobject]@{
        Status = "PASS"
        Mode = "VERIFY-ONLY"
        Version = $releaseInfo.Version
        Channel = $releaseInfo.Channel
        VstoInstaller = $vstoInstaller
        ExistingInstall = ($null -ne $existing)
        ExistingManifest = if ($null -eq $existing) { "" } else { [string]$existing.Manifest }
        ExcelOpen = [bool](Get-Process EXCEL -ErrorAction SilentlyContinue)
    } | Format-List
    return
}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Hay dong tat ca cua so Excel truoc khi cai dat hoac nang cap add-in."
}

if ($releaseInfo.Channel -eq "PILOT") {
    if (-not $TrustPilotCertificate) {
        throw "Ban PILOT dung certificate cuc bo. Chay lai voi -TrustPilotCertificate sau khi da kiem tra thumbprint trong release-info.json."
    }
    if (-not (Test-Path -LiteralPath $certificatePath)) { throw "Thieu public pilot certificate." }
    $certificate = New-Object Security.Cryptography.X509Certificates.X509Certificate2($certificatePath)
    if ($certificate.Thumbprint -ne $releaseInfo.PublisherThumbprint) {
        throw "Public certificate khong khop release-info.json."
    }
    Import-Certificate -FilePath $certificatePath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
    Import-Certificate -FilePath $certificatePath -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
}

if ($null -ne $existing -and -not [string]::IsNullOrWhiteSpace([string]$existing.Manifest)) {
    $oldManifestValue = [string]$existing.Manifest
    $oldManifest = Convert-ManifestValueToPath $oldManifestValue
    $isDevelopmentRegistration =
        $oldManifestValue.EndsWith("|vstolocal", [StringComparison]::OrdinalIgnoreCase) -and
        $oldManifest -match '[\\/]bin[\\/]'
    if ($isDevelopmentRegistration) {
        Remove-StaleAddinRegistration $existing "Visual Studio vstolocal registration"
    }
    elseif ([string]::IsNullOrWhiteSpace($oldManifest) -or -not (Test-Path -LiteralPath $oldManifest)) {
        Remove-StaleAddinRegistration $existing "Registered manifest no longer exists"
    }
    else {
        Write-Host "Dang go ban da cai tu: $oldManifest"
        $uninstalled = Invoke-VstoInstaller "/Uninstall" $oldManifest -AllowNotInstalled -ForceSilent
        if (-not $uninstalled -and (Test-Path -LiteralPath $addinRegistryPath)) {
            Remove-StaleAddinRegistration $existing "VSTOInstaller reported not installed (-401)"
        }
    }
}

# Visual Studio co the ghi de khoa Excel trong khi ClickOnce subscription cu van con.
Remove-InstalledProductSubscriptions ([string]$releaseInfo.Product)

Write-Host "Dang cai TTBMVN Excel Tools $($releaseInfo.Version)..."
try {
    [void](Invoke-VstoInstaller "/Install" $vstoPath)
}
catch {
    Restore-StaleAddinRegistration
    throw
}
$installed = Get-ItemProperty -LiteralPath $addinRegistryPath -ErrorAction Stop
if ([int]$installed.LoadBehavior -ne 3) { throw "Add-in da cai nhung LoadBehavior khong bang 3." }

$receiptRoot = Join-Path $env:LOCALAPPDATA "TTBMVNExcelTools"
New-Item -ItemType Directory -Path $receiptRoot -Force | Out-Null
$receipt = [ordered]@{
    Product = $releaseInfo.Product
    Version = $releaseInfo.Version
    Channel = $releaseInfo.Channel
    InstalledAtUtc = [DateTime]::UtcNow.ToString("o")
    ReleaseArchiveSha256 = $releaseInfo.ReleaseArchiveSha256
    PublisherThumbprint = $releaseInfo.PublisherThumbprint
}
[IO.File]::WriteAllText(
    (Join-Path $receiptRoot "install-receipt.json"),
    ($receipt | ConvertTo-Json -Depth 4),
    (New-Object Text.UTF8Encoding($false)))
Write-Host "Cai dat thanh cong. Mo lai Excel de nap add-in."
