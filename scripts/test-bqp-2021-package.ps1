param(
    [string]$WorkbookPath = "Dutoanmau\00. Du toan TP 3.xlsm",
    [switch]$SkipWorkbook
)

$ErrorActionPreference = "Stop"

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    if ($Expected -ne $Actual) {
        throw "$Message Expected=[$Expected] Actual=[$Actual]"
    }
}

function Test-UsedRangeContains {
    param($Worksheet, [string]$Needle)
    $usedRange = $Worksheet.UsedRange
    try {
        $values = $usedRange.Value2
        if ($usedRange.Rows.Count -eq 1 -and $usedRange.Columns.Count -eq 1) {
            return ([string]$values).IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0
        }
        for ($row = 1; $row -le $usedRange.Rows.Count; $row++) {
            for ($column = 1; $column -le $usedRange.Columns.Count; $column++) {
                $value = $values[$row, $column]
                if ($null -ne $value -and
                    ([string]$value).IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    return $true
                }
            }
        }
        return $false
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange)
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$packageRoot = Join-Path $repoRoot "data\regulations\packages\BQP-RPBM-2021"
$sourceRoot = Join-Path $packageRoot "source"
$bundleRoot = Join-Path $packageRoot "bundle"
$tool = Join-Path $repoRoot "ExcelAddIn1.RegulationTool\bin\Release\ExcelAddIn1.RegulationTool.exe"

Assert-True (Test-Path -LiteralPath $tool) "RegulationTool Release chua duoc build."
& $tool validate $bundleRoot
if ($LASTEXITCODE -ne 0) {
    throw "Bundle validation failed."
}

$sourceDocuments = Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $sourceRoot "sources.tsv")
$expectedDocuments = @{
    "TT121-2021-BQP" = @("TT121-2021-BQP.pdf", "9C0C9870EFD49FA49F7CC345B9D6D538D62AB5917EC9B68AE0B0C1E167363BA3")
    "TT122-2021-BQP" = @("TT122-2021-BQP.pdf", "C5F78850F99411518267445B83F97124843666EEECDE06D5946C3DC24D0407EC")
    "TT123-2021-BQP" = @("TT123-2021-BQP.pdf", "DE436E3BE81CE8B62A089D5B36C8A7A931BDB9EAC468D65B110455CB1F0DCF2E")
}
Assert-Equal 3 $sourceDocuments.Count "Sai so van ban nguon."
foreach ($document in $sourceDocuments) {
    Assert-True $expectedDocuments.ContainsKey($document.documentId) (
        "Van ban nguon khong mong doi: " + $document.documentId)
    $expected = $expectedDocuments[$document.documentId]
    $pdf = Join-Path $repoRoot ("data\regulations\raw\2021\" + $expected[0])
    $actualHash = (Get-FileHash -LiteralPath $pdf -Algorithm SHA256).Hash
    Assert-Equal $expected[1] $actualHash ("Sai SHA-256 PDF " + $document.documentId + ".")
    Assert-Equal $expected[1] $document.contentChecksum (
        "Checksum source manifest sai " + $document.documentId + ".")
}

$expectedCounts = [ordered]@{
    TechnicalProcess = 59
    Norm = 34
    CostRule = 34
    MachineRate = 33
    Geography = 35
    Compliance = 3
}
$allKeys = @{}
foreach ($entry in $expectedCounts.GetEnumerator()) {
    $path = Join-Path $sourceRoot ("modules\" + $entry.Key + ".tsv")
    $rows = @(Import-Csv -Delimiter "`t" -Encoding UTF8 $path)
    Assert-Equal $entry.Value $rows.Count ("Sai record count " + $entry.Key + ".")
    foreach ($row in $rows) {
        Assert-True (-not [string]::IsNullOrWhiteSpace($row.key)) ("Key trong " + $entry.Key + ".")
        Assert-True (-not $allKeys.ContainsKey($row.key)) ("Trung key " + $row.key + ".")
        $allKeys.Add($row.key, $entry.Key)
        Assert-Equal "VerifiedAgainstOfficialSource" $row.verification (
            "Record chua xac minh " + $row.key + ".")
        Assert-True ([int]$row.pageFrom -ge 1 -and [int]$row.pageTo -ge [int]$row.pageFrom) (
            "Locator trang sai " + $row.key + ".")
        Assert-True (-not [string]::IsNullOrWhiteSpace($row.data)) ("Data trong " + $row.key + ".")
    }
}
Assert-Equal 198 $allKeys.Count "Sai tong record package 2021."

# Review pass A: values read directly from official PDF tables.
$machineRows = Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $sourceRoot "modules\MachineRate.tsv")
$machine004 = $machineRows | Where-Object key -eq "MACHINE-M010.004"
$machine025 = $machineRows | Where-Object key -eq "MACHINE-M010.025"
Assert-True ($machine004.data -match "annualShifts=280") "M010.004 sai so ca nam."
Assert-True ($machine004.data -match "referencePriceVnd=566835000") "M010.004 sai nguyen gia."
Assert-True ($machine025.data -match "referencePriceVnd=890000") "M010.025 sai nguyen gia."

$costRows = Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $sourceRoot "modules\CostRule.tsv")
$commonCost = $costRows | Where-Object key -eq "COST-COMMON"
$civilK5 = $costRows | Where-Object key -eq "COST-K5-CIVIL"
Assert-True ($commonCost.data -match "ratePercent=40;base=NC") "Chi phi chung sai can cu."
Assert-True ($civilK5.data -match "ratesPercent=3.285,2.853,2.435") "Bang K5 dan dung sai."

if (-not $SkipWorkbook) {
    # Review pass B: independent oracle in the canonical workbook.
    $expectedWorkbook = [IO.Path]::GetFullPath((Join-Path $repoRoot $WorkbookPath))
    $openSession = Get-OpenExcelWorkbookSession $expectedWorkbook
    Assert-True ($null -ne $openSession) "Workbook chuan chua mo trong Excel."
    $excel = $openSession.Application
    $workbook = $openSession.Workbook
    try {
        $lookup = $workbook.Worksheets.Item("Tracuu")
        $cost = $workbook.Worksheets.Item("ChiPhi")
        $land = $workbook.Worksheets.Item("DG Can")
        $water = $workbook.Worksheets.Item("DG Nuoc")
        try {
            Assert-Equal 245 ([double]$lookup.Range("C22").Value2) "Tracuu khu vuc 4, do sau 0.3/0.5 m sai."
            Assert-Equal 25 ([double]$lookup.Range("D22").Value2) "Tracuu khu vuc 4, do sau den 1 m sai."
            Assert-Equal 6 ([double]$lookup.Range("E22").Value2) "Tracuu khu vuc 4, do sau den 3 m sai."
            Assert-Equal 1.5 ([double]$lookup.Range("F22").Value2) "Tracuu khu vuc 4, do sau den 5 m sai."
            Assert-Equal 0.2 ([double]$lookup.Range("G22").Value2) "Tracuu khu vuc 4, do sau den 10 m sai."
            Assert-Equal 3.285 ([double]$cost.Range("D8").Value2) "ChiPhi K5 dan dung sai."

            Assert-True (Test-UsedRangeContains $land "020.0800") "DG Can thieu ma 020.0800."
            Assert-True (Test-UsedRangeContains $water "030.0700") "DG Nuoc thieu ma 030.0700."
        }
        finally {
            foreach ($sheet in @($lookup, $cost, $land, $water)) {
                if ($null -ne $sheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
            }
        }
    }
    finally {
        Release-OpenExcelWorkbookSession $openSession
    }
}

Write-Host "BQP 2021 package validation passed: 198 records; PDF and workbook reviews passed."
