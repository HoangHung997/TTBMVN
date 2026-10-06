param(
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\00. Du toan TP 3.xlsm"
}
$expectedPath = [IO.Path]::GetFullPath($WorkbookPath)
$allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau"))
if (-not $expectedPath.StartsWith(
    $allowedRoot.TrimEnd('\') + '\',
    [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook test nam ngoai Dutoanmau: $expectedPath"
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    if ($Expected -ne $Actual) {
        throw "$Message Expected=[$Expected] Actual=[$Actual]"
    }
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Find-Workbook([object]$Excel, [string]$FullPath) {
    $workbooks = $Excel.Workbooks
    try {
        for ($index = 1; $index -le $workbooks.Count; $index++) {
            $candidate = $workbooks.Item($index)
            if ([string]::Equals(
                [IO.Path]::GetFullPath($candidate.FullName),
                $FullPath,
                [StringComparison]::OrdinalIgnoreCase)) {
                return $candidate
            }
            Release-ComObject $candidate
        }
    }
    finally {
        Release-ComObject $workbooks
    }
    return $null
}

function Get-CellProperty([object]$Sheet, [string]$Address, [string]$Property) {
    $cell = $Sheet.Range($Address)
    try {
        return $cell.$Property
    }
    finally {
        Release-ComObject $cell
    }
}

$bundle = Join-Path $repoRoot "data\regulations\packages\BQP-RPBM-2025\bundle"
$tool = Join-Path $repoRoot "ExcelAddIn1.RegulationTool\bin\Release\ExcelAddIn1.RegulationTool.exe"
$tests = Join-Path $repoRoot "ExcelAddIn1.Tests\bin\Release\ExcelAddIn1.Tests.exe"
& $tool validate $bundle | Out-Host
if ($LASTEXITCODE -ne 0) { throw "Validate bundle 2025 failed." }
$testOutput = @(& $tests 2>&1)
if ($LASTEXITCODE -ne 0) {
    $testOutput | Out-Host
    throw "Core tests failed."
}
foreach ($name in @("MachineRateCatalog", "MachineRateCalculation", "MachineRateLocaleAndValidation")) {
    if ($testOutput -notcontains "[PASS] $name") {
        throw "Thieu ket qua test $name."
    }
}

$expected = [ordered]@{
    "F23" = 762996
    "F29" = 1070324
    "F35" = 1587824
    "F41" = 518226
    "F46" = 554101
    "F51" = 995354
    "F56" = 2946909
    "F62" = 1105582
    "F67" = 1108263
    "F72" = 1112914
}

$openSession = Get-OpenExcelWorkbookSession $expectedPath
if ($null -eq $openSession) {
    throw "Workbook chuan chua mo trong Excel: $expectedPath"
}
$excel = $openSession.Application
$workbook = $openSession.Workbook
$sheet = $null
try {
    $sheet = $workbook.Worksheets.Item("VL-NC-M")
    foreach ($entry in $expected.GetEnumerator()) {
        $cell = $sheet.Range($entry.Key)
        try {
            Assert-Equal ([double]$entry.Value) ([double]$cell.Value2) (
                "VL-NC-M!" + $entry.Key + " sai.")
        }
        finally {
            Release-ComObject $cell
        }
    }
    Assert-Equal "=SUM(F52:F55)" ([string](Get-CellProperty $sheet "F51" "Formula")) `
        "Thuyen composite phai loai nhien lieu khoi gia ca may trong workbook chuan."
    Assert-Equal "=SUM(F57:F61)" ([string](Get-CellProperty $sheet "F56" "Formula")) `
        "May hut/xoi bun cat phai gom nhien lieu trong workbook chuan."
    Assert-Equal 517500 ([double](Get-CellProperty $sheet "F15" "Value2")) `
        "Gia nhan cong bac 8/10 sai."
}
finally {
    Release-ComObject $sheet
    Release-OpenExcelWorkbookSession $openSession
}

Write-Host (
    "Machine rate validation passed: 33 machines; 3 core gates; " +
    $expected.Count + " workbook totals.")
