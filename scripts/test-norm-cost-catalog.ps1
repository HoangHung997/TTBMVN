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

function Assert-Near {
    param([double]$Expected, [double]$Actual, [double]$Tolerance, [string]$Message)
    if ([Math]::Abs($Expected - $Actual) -gt $Tolerance) {
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

function Get-CellValue([object]$Sheet, [string]$Address) {
    $cell = $Sheet.Range($Address)
    try { return $cell.Value2 }
    finally { Release-ComObject $cell }
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
foreach ($name in @(
    "NormCatalog",
    "NormCalculation",
    "NormConditionsAndValidation",
    "CostRuleCatalog",
    "CostRuleCalculation",
    "CostRuleLocaleAndValidation")) {
    if ($testOutput -notcontains "[PASS] $name") {
        throw "Thieu ket qua test $name."
    }
}

$officialFiles = [ordered]@{
    "TT36-2026-BXD.pdf" = "9B1C02325E76F90261421F49648FE8951362BF29C48A79377ED45F5121609923"
    "TT36-2026-BXD-PL3.pdf" = "936E5C591A85C8EDE5D0375E23FF3EA5A96C1124834F2162FFC80AEC484CBD9F"
    "TT38-2026-BXD.pdf" = "604543B5F13EBDC529C9DD4C6D7537642FF1A462499305B7DFC05AD3F1E28F30"
    "TT38-2026-BXD-PL8.pdf" = "74738C35C1DA28E78CF78903C280EFB69D287295C7F88C49B70F75C3D0D9FBBF"
}
foreach ($entry in $officialFiles.GetEnumerator()) {
    $path = Join-Path $repoRoot ("data\regulations\raw\2026\" + $entry.Key)
    Assert-Equal $entry.Value (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash (
        "Sai checksum nguon " + $entry.Key + ".")
}

$openSession = Get-OpenExcelWorkbookSession $expectedPath
if ($null -eq $openSession) {
    throw "Workbook chuan chua mo trong Excel: $expectedPath"
}
$excel = $openSession.Application
$workbook = $openSession.Workbook
$land = $null
$cost = $null
try {
    $land = $workbook.Worksheets.Item("DG Can")
    $normExpected = [ordered]@{
        "D85" = 50
        "D86" = 100
        "D87" = 1
        "D88" = 5
        "D90" = 7.05
        "D92" = 4.70
        "D119" = 0.2
        "D120" = 0.004
        "D121" = 0.15
        "D122" = 1
        "D124" = 1.14
        "D126" = 0.008
    }
    foreach ($entry in $normExpected.GetEnumerator()) {
        Assert-Near ([double]$entry.Value) ([double](Get-CellValue $land $entry.Key)) 0.0000001 (
            "DG Can!" + $entry.Key + " sai hao phi.")
    }

    $cost = $workbook.Worksheets.Item("ChiPhi")
    $k5Expected = @(
        @(3.285, 2.853, 2.435, 1.845, 1.546, 1.188, 0.797, 0.694, 0.620, 0.530, 0.478),
        @(3.508, 3.137, 2.559, 2.074, 1.604, 1.301, 0.823, 0.716, 0.640, 0.550, 0.493),
        @(3.203, 2.700, 2.356, 1.714, 1.272, 1.003, 0.731, 0.636, 0.550, 0.480, 0.438),
        @(2.598, 2.292, 2.075, 1.545, 1.189, 0.950, 0.631, 0.550, 0.490, 0.420, 0.378),
        @(2.566, 2.256, 1.984, 1.461, 1.142, 0.912, 0.584, 0.509, 0.452, 0.390, 0.350)
    )
    for ($row = 0; $row -lt $k5Expected.Count; $row++) {
        for ($column = 0; $column -lt $k5Expected[$row].Count; $column++) {
            $cell = $cost.Cells.Item(8 + $row, 4 + $column)
            try {
                Assert-Near ([double]$k5Expected[$row][$column]) ([double]$cell.Value2) 0.0000001 (
                    "ChiPhi K5 row=" + $row + " column=" + $column + " sai.")
            }
            finally { Release-ComObject $cell }
        }
    }
    foreach ($entry in ([ordered]@{
        "D22" = 2.2; "E22" = 2.0; "F22" = 1.9; "G22" = 1.8; "H22" = 1.7
        "D23" = 1.1; "E23" = 1.0; "F23" = 0.95; "G23" = 0.9; "H23" = 0.85
    }).GetEnumerator()) {
        Assert-Near ([double]$entry.Value) ([double](Get-CellValue $cost $entry.Key)) 0.0000001 (
            "ChiPhi!" + $entry.Key + " sai K2.")
    }
}
finally {
    Release-ComObject $cost
    Release-ComObject $land
    Release-OpenExcelWorkbookSession $openSession
}

Write-Host "Norm/cost validation passed: 6 core gates; 12 norm cells; 55 K5 cells; 10 K2 cells; 4 official PDFs."
