param(
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-505-start.xlsm"
}
$workbookPath = [IO.Path]::GetFullPath($WorkbookPath)
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)

$loggerType = $addin.GetType("ExcelAddIn1.Funtion.RuntimeLogger", $true)
$supportType = $addin.GetType("ExcelAddIn1.Funtion.SupportPackage", $true)
$beginOperation = $loggerType.GetMethods() | Where-Object {
    $_.Name -eq "BeginOperation" -and $_.GetParameters().Length -eq 4
} | Select-Object -First 1
$createSupport = $supportType.GetMethods() | Where-Object {
    $_.Name -eq "Create" -and $_.GetParameters().Length -eq 1
} | Select-Object -First 1
$getWorkbookToken = $loggerType.GetMethod("GetWorkbookToken")
$logPath = [ExcelAddIn1.Funtion.RuntimeLogger]::LogFilePath
$initialLogLength = if (Test-Path -LiteralPath $logPath) {
    (Get-Item -LiteralPath $logPath).Length
}
else { 0L }

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

function Assert-DoesNotContain([string]$Text, [string]$Sensitive, [string]$Label) {
    if (-not [string]::IsNullOrEmpty($Sensitive) -and
        $Text.IndexOf($Sensitive, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
        throw "$Label con du lieu nhay cam: $Sensitive"
    }
}

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
    $machineId = [ExcelAddIn1.Funtion.LicenseManager]::GetMachineId()
    $workbookName = $workbook.Name
    $workbookFullName = $workbook.FullName
    $fakeLicense = "ABCDE-FGHIJ-KLMNO-PQRST-UVWXY"
    $secretValue = "SENSITIVE_CUSTOMER_VALUE_DT505"
    $scope = Invoke-Static $beginOperation @(
        "DT-505", "controlled-com", $workbook, "EstimateAppendix")
    $correlationId = $scope.CorrelationId
    try {
        $exception = [System.Runtime.InteropServices.COMException]::new(
            "Controlled HRESULT at $workbookFullName with key $fakeLicense",
            [int]-2146827284)
        $context = New-Object "System.Collections.Generic.Dictionary[string,string]"
        $context.Add("SafeFlag", "controlled")
        $context.Add("CustomerName", $secretValue)
        $context.Add("LicenseKey", $fakeLicense)
        $context.Add("TableDataAddress", "'Secret Sheet'!A1:B2")
        [ExcelAddIn1.Funtion.RuntimeLogger]::Log(
            $exception,
            "Controlled COM diagnostic test",
            $context)
    }
    finally { $scope.Dispose() }

    $allLog = [IO.File]::ReadAllText($logPath, [Text.Encoding]::UTF8)
    $logRecord = if ($initialLogLength -lt (Get-Item -LiteralPath $logPath).Length) {
        $stream = [IO.File]::Open($logPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            [void]$stream.Seek($initialLogLength, [IO.SeekOrigin]::Begin)
            $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::UTF8)
            try { $reader.ReadToEnd() }
            finally { $reader.Dispose() }
        }
        finally { $stream.Dispose() }
    }
    else { throw "Runtime log khong duoc ghi." }

    $required = @(
        "CorrelationId: $correlationId",
        "Task: DT-505",
        "Phase: controlled-com",
        "WorkbookId: WB-",
        "WorksheetRole: EstimateAppendix",
        "PackageId: BQP-RPBM-2025",
        "PackageVersion: 2.0.1",
        "Exception.Type: System.Runtime.InteropServices.COMException",
        "Exception.HResult: 0x800A03EC",
        "Context.SafeFlag: controlled",
        "Context.CustomerName: <redacted>",
        "Context.LicenseKey: <redacted>",
        "Context.TableDataAddress: <redacted>"
    )
    foreach ($value in $required) {
        if ($logRecord.IndexOf($value, [StringComparison]::Ordinal) -lt 0) {
            throw "Runtime log thieu: $value"
        }
    }
    Assert-DoesNotContain $logRecord $workbookName "Runtime log"
    Assert-DoesNotContain $logRecord $workbookFullName "Runtime log"
    Assert-DoesNotContain $logRecord $machineId "Runtime log"
    Assert-DoesNotContain $logRecord $fakeLicense "Runtime log"
    Assert-DoesNotContain $logRecord $secretValue "Runtime log"

    $supportFolder = Invoke-Static $createSupport @($workbook)
    $expectedFiles = @(
        "app-info.txt",
        "activation-info.txt",
        "workbook-info.txt",
        "project-profile.txt",
        "package-manifest.ttbmanifest",
        "validation-report.txt",
        "daodat-settings.json",
        "runtime.log",
        "manifest.sha256"
    )
    foreach ($fileName in $expectedFiles) {
        if (-not (Test-Path -LiteralPath (Join-Path $supportFolder $fileName))) {
            throw "Support package thieu $fileName."
        }
    }
    $supportText = ($expectedFiles | Where-Object { $_ -ne "manifest.sha256" } | ForEach-Object {
        [IO.File]::ReadAllText((Join-Path $supportFolder $_), [Text.Encoding]::UTF8)
    }) -join "`n"
    foreach ($pair in @(
        @($workbookName, "workbook name"),
        @($workbookFullName, "workbook path"),
        @($machineId, "machine id"),
        @($fakeLicense, "license key"),
        @($secretValue, "customer value"),
        @("Secret Sheet", "sheet name")
    )) {
        Assert-DoesNotContain $supportText $pair[0] ("Support package " + $pair[1])
    }
    if ($supportText.IndexOf("MachineToken: M-", [StringComparison]::Ordinal) -lt 0 -or
        $supportText.IndexOf("WorkbookId: WB-", [StringComparison]::Ordinal) -lt 0 -or
        $supportText.IndexOf("PackageId: BQP-RPBM-2025", [StringComparison]::Ordinal) -lt 0 -or
        $supportText.IndexOf("ProjectToken: PRJ-", [StringComparison]::Ordinal) -lt 0 -or
        $supportText.IndexOf("TTBMVN_REGULATION_PACKAGE", [StringComparison]::Ordinal) -lt 0 -or
        $supportText.IndexOf("Valid: True", [StringComparison]::Ordinal) -lt 0 -or
        $supportText.IndexOf('"IncludeWorkbookFields":false', [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw "Support package thieu token/package hoac setting chua duoc loc."
    }
    $manifest = Get-Content -LiteralPath (Join-Path $supportFolder "manifest.sha256")
    if ($manifest.Count -ne 8) {
        throw "Support manifest phai co 8 file payload, actual=$($manifest.Count)."
    }

    [pscustomobject]@{
        CorrelationId = $correlationId
        Task = "DT-505"
        Phase = "controlled-com"
        HResult = "0x800A03EC"
        WorkbookToken = Invoke-Static $getWorkbookToken @($workbook)
        MachineToken = [ExcelAddIn1.Funtion.RuntimeLogger]::GetMachineToken()
        Package = "BQP-RPBM-2025@2.0.1"
        ContextRedaction = "PASS"
        SupportFiles = $expectedFiles.Count
        Manifest = "PASS"
        SupportFolder = $supportFolder
    } | Format-List
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}
