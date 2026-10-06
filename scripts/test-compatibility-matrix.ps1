param(
    [string]$CheckpointPath = ".\Dutoanmau\Checkpoints\DT-603-start.xlsm"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$buildScript = Join-Path $PSScriptRoot "build-release.ps1"
$checkpoint = [IO.Path]::GetFullPath((Join-Path $repoRoot $CheckpointPath))
$evidencePath = Join-Path $repoRoot "tmp\DT-603-compatibility-current.json"

& $buildScript -Platform "Any CPU"
if ($LASTEXITCODE -ne 0) { throw "Release|Any CPU failed." }
& $buildScript -Platform "x64"
if ($LASTEXITCODE -ne 0) { throw "Release|x64 failed." }

$os = Get-CimInstance Win32_OperatingSystem
$computer = Get-CimInstance Win32_ComputerSystem
$officePath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$office = Get-ItemProperty $officePath -ErrorAction Stop
$dotNet = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction Stop
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$cultureType = $addin.GetType("ExcelAddIn1.Funtion.ExcelCulture", $true)
$getCulture = $cultureType.GetMethod("GetNumberCulture")

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

function Test-ExcelSeparatorScenario(
    [object]$Excel,
    [object]$Sheet,
    [string]$Decimal,
    [string]$Thousands,
    [string]$ExpectedText) {
    $Excel.UseSystemSeparators = $false
    $Excel.DecimalSeparator = $Decimal
    $Excel.ThousandsSeparator = $Thousands
    $culture = Invoke-Static $getCulture @($Excel)
    if ($culture.NumberFormat.NumberDecimalSeparator -ne $Decimal -or
        $culture.NumberFormat.NumberGroupSeparator -ne $Thousands) {
        throw "ExcelCulture khong theo separator Excel: decimal=$Decimal thousands=$Thousands"
    }
    $formatted = ([decimal]1234.5).ToString("N2", $culture)
    if ($formatted -ne $ExpectedText) {
        throw "Format separator sai. Expected=$ExpectedText Actual=$formatted"
    }

    $Sheet.Range("A1").Value2 = 1.25
    $Sheet.Range("B1").Value2 = 2.5
    $Sheet.Range("C1").FormulaR1C1 = "=SUM(RC[-2],RC[-1])"
    $Sheet.Range("D1").FormulaR1C1 = "=ROUND(RC[-1],2)"
    $Excel.Calculate()
    if ([math]::Abs([double]$Sheet.Range("D1").Value2 - 3.75) -gt 0.0000001) {
        throw "FormulaR1C1 invariant sai voi separator $Decimal/$Thousands."
    }
    return [pscustomobject]@{
        Decimal = $Decimal
        Thousands = $Thousands
        List = [string]$Excel.International(5)
        Formatted = $formatted
        FormulaValue = [double]$Sheet.Range("D1").Value2
        Status = "PASS"
    }
}

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$scratch = $null
$sheet = $null
$workbook = $null
$oldUseSystem = $null
$oldDecimal = $null
$oldThousands = $null
try {
    $oldUseSystem = $excel.UseSystemSeparators
    $oldDecimal = $excel.DecimalSeparator
    $oldThousands = $excel.ThousandsSeparator
    $defaultSeparators = [pscustomobject]@{
        Decimal = [string]$excel.International(3)
        Thousands = [string]$excel.International(4)
        List = [string]$excel.International(5)
    }

    $scratch = $excel.Workbooks.Add()
    $sheet = $scratch.Worksheets.Item(1)
    $commaDecimal = Test-ExcelSeparatorScenario $excel $sheet "," "." "1.234,50"
    $dotDecimal = Test-ExcelSeparatorScenario $excel $sheet "." "," "1,234.50"

    $excel.UseSystemSeparators = $oldUseSystem
    $excel.DecimalSeparator = $oldDecimal
    $excel.ThousandsSeparator = $oldThousands
    $scratch.Close($false)
    Release-ComObject $sheet
    Release-ComObject $scratch
    $sheet = $null
    $scratch = $null

    $workbook = $excel.Workbooks.Open($checkpoint, 0, $true)
    $summary = $workbook.Worksheets.Item("THKP-TC")
    try {
        $e27 = [double]$summary.Range("E27").Value2
        if ($e27 -ne 1201557000) { throw "Checkpoint E27 sai trong compatibility test." }
    }
    finally { Release-ComObject $summary }

    $evidence = [ordered]@{
        CapturedUtc = [DateTime]::UtcNow.ToString("O", [Globalization.CultureInfo]::InvariantCulture)
        OS = [ordered]@{
            Caption = $os.Caption
            Version = $os.Version
            Build = $os.BuildNumber
            Architecture = $os.OSArchitecture
            SystemType = $computer.SystemType
        }
        Office = [ordered]@{
            Product = $office.ProductReleaseIds
            Version = $office.VersionToReport
            Platform = $office.Platform
            ClientCulture = $office.ClientCulture
            ExcelVersion = $excel.Version
            ExcelBuild = $excel.Build
        }
        Runtime = [ordered]@{
            DotNetRelease = $dotNet.Release
            ProcessCulture = [Globalization.CultureInfo]::CurrentCulture.Name
            ProcessUICulture = [Globalization.CultureInfo]::CurrentUICulture.Name
        }
        Build = [ordered]@{
            ReleaseAnyCPU = "PASS"
            ReleaseX64 = "PASS"
            CoreTestsEachBuild = 68
        }
        ExcelDefaultSeparators = $defaultSeparators
        SeparatorScenarios = @($commaDecimal, $dotDecimal)
        Checkpoint = [ordered]@{
            Path = $checkpoint
            E27 = $e27
            Status = "PASS"
        }
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $evidencePath) -Force | Out-Null
    $evidence | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $evidencePath -Encoding UTF8
    $evidence | ConvertTo-Json -Depth 6
}
finally {
    if ($null -ne $excel -and $null -ne $oldUseSystem) {
        try {
            $excel.UseSystemSeparators = $oldUseSystem
            $excel.DecimalSeparator = $oldDecimal
            $excel.ThousandsSeparator = $oldThousands
        }
        catch { }
    }
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $scratch) { $scratch.Close($false) }
    Release-ComObject $sheet
    Release-ComObject $scratch
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}

Write-Host "DT-603 current-host compatibility PASS."
