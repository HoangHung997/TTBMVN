$script:TtbmvnPilotSigningSubject = "CN=TTBMVN Pilot Publisher"
$script:TtbmvnCodeSigningEku = "1.3.6.1.5.5.7.3.3"

function Test-TtbmvnCodeSigningCertificate([object]$Certificate) {
    if ($null -eq $Certificate -or -not $Certificate.HasPrivateKey) { return $false }
    $now = Get-Date
    if ($Certificate.NotBefore -gt $now -or $Certificate.NotAfter -le $now.AddDays(30)) { return $false }
    $eku = $Certificate.Extensions | Where-Object { $_.Oid.Value -eq "2.5.29.37" } | Select-Object -First 1
    return $null -ne $eku -and $eku.Format($false).Contains($script:TtbmvnCodeSigningEku)
}

function Find-TtbmvnCertificateByThumbprint([string]$Thumbprint) {
    $normalized = if ($null -eq $Thumbprint) { "" } else { $Thumbprint }
    $normalized = $normalized.Replace(" ", "").ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $null }
    foreach ($store in @("Cert:\CurrentUser\My", "Cert:\LocalMachine\My")) {
        $certificate = Get-ChildItem -LiteralPath $store -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $normalized } |
            Select-Object -First 1
        if ($null -ne $certificate) { return $certificate }
    }
    return $null
}

function Get-TtbmvnManifestSigningCertificate {
    if (-not [string]::IsNullOrWhiteSpace($env:TTBMVN_SIGNING_CERT_THUMBPRINT)) {
        $configured = Find-TtbmvnCertificateByThumbprint $env:TTBMVN_SIGNING_CERT_THUMBPRINT
        if (-not (Test-TtbmvnCodeSigningCertificate $configured)) {
            throw "TTBMVN_SIGNING_CERT_THUMBPRINT khong tro den Code Signing certificate hop le co private key."
        }
        return $configured
    }

    $pilot = Get-ChildItem Cert:\CurrentUser\My -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -eq $script:TtbmvnPilotSigningSubject } |
        Where-Object { Test-TtbmvnCodeSigningCertificate $_ } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1
    if ($null -eq $pilot) {
        $pilot = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject $script:TtbmvnPilotSigningSubject `
            -FriendlyName "TTBMVN Excel Tools Pilot Manifest Signing" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -KeyAlgorithm RSA `
            -KeyLength 3072 `
            -HashAlgorithm SHA256 `
            -KeyExportPolicy NonExportable `
            -NotAfter (Get-Date).AddYears(5)
    }
    if (-not (Test-TtbmvnCodeSigningCertificate $pilot)) {
        throw "Khong tao duoc pilot Code Signing certificate hop le."
    }
    return $pilot
}

function Get-TtbmvnSigningChannel([object]$Certificate) {
    if ($Certificate.Subject -eq $script:TtbmvnPilotSigningSubject) { return "PILOT" }
    return "EXTERNAL"
}
