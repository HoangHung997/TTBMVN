param(
    [switch]$RegenerateSource
)

$ErrorActionPreference = "Stop"

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    if ($Expected -ne $Actual) {
        throw "$Message Expected=[$Expected] Actual=[$Actual]"
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$packageRoot = Join-Path $repoRoot "data\regulations\packages\BQP-RPBM-2025"
$sourceRoot = Join-Path $packageRoot "source"
$bundleRoot = Join-Path $packageRoot "bundle"
$tool = Join-Path $repoRoot "ExcelAddIn1.RegulationTool\bin\Release\ExcelAddIn1.RegulationTool.exe"

if ($RegenerateSource) {
    $python = "C:\Users\hoang\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
    Assert-True (Test-Path -LiteralPath $python) "Khong tim thay Python runtime de regenerate source."
    & $python (Join-Path $repoRoot "tools\regulations\build_bqp_2025_source.py") `
        --repository-root $repoRoot
    if ($LASTEXITCODE -ne 0) { throw "Regenerate source 2025 failed." }
}

Assert-True (Test-Path -LiteralPath $tool) "RegulationTool Release chua duoc build."
& $tool validate $bundleRoot
if ($LASTEXITCODE -ne 0) { throw "Bundle validation failed." }

$expectedDocuments = [ordered]@{
    "TT101-2025-BQP" = @("TT101-2025-BQP.pdf", "E445F63CFDB515CB187780247423610A5A29DD0BBFDB231EBC21D858516DD240", 69)
    "VBHN94-2025-BQP" = @("VBHN94-2025-BQP.pdf", "D375AAC8F0C0B016D0C8B1E8CC65140560DF23D13AB3B499DE581DC5276F8935", 30)
    "VBHN95-2025-BQP" = @("VBHN95-2025-BQP.pdf", "2D5207B562EC63A24BD0B9D38C86E748AF6F7A4009195250517B83EFE22A1094", 99)
    "VBHN96-2025-BQP" = @("VBHN96-2025-BQP.pdf", "10A25B90BC95F8EB8330C9CD6D2423AD0DF46746D95DB7D33C7A2157228AE0B1", 35)
    "VBHN97-2025-BQP" = @("VBHN97-2025-BQP.pdf", "92B6C42996E554A908661B0395CE7286D41DDD69C17859B40993DC46300B69E0", 52)
    "VBHN98-2025-BQP" = @("VBHN98-2025-BQP.pdf", "62CC32BDA9ADED4310AB183180D6D4D861E41FEA94852848425EA0C19BAFE0A2", 76)
}
$sourceDocuments = @(Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $sourceRoot "sources.tsv"))
$textManifest = @(Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $repoRoot "data\regulations\text\2025\text-manifest.tsv"))
Assert-Equal 6 $sourceDocuments.Count "Sai so van ban nguon 2025."
Assert-Equal 6 $textManifest.Count "Sai so van ban trong text manifest 2025."
foreach ($document in $sourceDocuments) {
    Assert-True $expectedDocuments.Contains($document.documentId) (
        "Van ban nguon khong mong doi: " + $document.documentId)
    $expected = $expectedDocuments[$document.documentId]
    $pdf = Join-Path $repoRoot ("data\regulations\raw\2025\" + $expected[0])
    $actualHash = (Get-FileHash -LiteralPath $pdf -Algorithm SHA256).Hash
    Assert-Equal $expected[1] $actualHash ("Sai SHA-256 PDF " + $document.documentId + ".")
    Assert-Equal $expected[1] $document.contentChecksum (
        "Checksum source manifest sai " + $document.documentId + ".")
    $manifestRow = $textManifest | Where-Object fileName -eq $expected[0]
    Assert-True ($null -ne $manifestRow) ("Text manifest thieu " + $expected[0] + ".")
    Assert-Equal $expected[1] $manifestRow.pdfSha256 ("Text manifest sai PDF hash " + $expected[0] + ".")
    Assert-Equal $expected[2] ([int]$manifestRow.pageCount) ("Text manifest sai so trang " + $expected[0] + ".")
}

$expectedCounts = [ordered]@{
    TechnicalProcess = 60
    Norm = 42
    CostRule = 37
    MachineRate = 33
    Geography = 69
    Compliance = 7
}
$allKeys = @{}
foreach ($entry in $expectedCounts.GetEnumerator()) {
    $rows = @(Import-Csv -Delimiter "`t" -Encoding UTF8 (
        Join-Path $sourceRoot ("modules\" + $entry.Key + ".tsv")))
    Assert-Equal $entry.Value $rows.Count ("Sai record count " + $entry.Key + ".")
    foreach ($row in $rows) {
        Assert-True (-not $allKeys.ContainsKey($row.key)) ("Trung key " + $row.key + ".")
        $allKeys.Add($row.key, $entry.Key)
        Assert-Equal "VerifiedAgainstOfficialSource" $row.verification (
            "Record chua xac minh " + $row.key + ".")
        Assert-True ([int]$row.pageFrom -ge 1 -and [int]$row.pageTo -ge [int]$row.pageFrom) (
            "Locator trang sai " + $row.key + ".")
        Assert-True (-not [string]::IsNullOrWhiteSpace($row.data)) ("Data trong " + $row.key + ".")
    }
}
Assert-Equal 248 $allKeys.Count "Sai tong record package 2025."

$norms = @(Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $sourceRoot "modules\Norm.tsv"))
$norm0200500 = $norms | Where-Object key -eq "NORM-020.0500"
Assert-True ($norm0200500.data -match "LAB-QNCN-7:worker-day:6.40,7.05,7.76,8.54") `
    "NORM-020.0500 sai hao phi nhan cong."
Assert-True ($norm0200500.data -match "M010.003:shift:0,0,0,5.69") `
    "NORM-020.0500 sai hao phi Vet1."
Assert-True ($norm0200500.data -match "MAT-RED-FLAG-LARGE:each:1,1,1,1") `
    "NORM-020.0500 chua tach co do 0,4x0,6m."

$machines = @(Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $sourceRoot "modules\MachineRate.tsv"))
Assert-True (($machines | Where-Object key -eq "MACHINE-M010.004").data -match `
    "referencePriceVnd=613644600") "M010.004 sai nguyen gia 2025."
Assert-True (($machines | Where-Object key -eq "MACHINE-M010.027").data -match `
    "fuel=60-lit-gasoline-E5-RON-92-II") "M010.027 thieu nhien lieu 2025."
Assert-True (($machines | Where-Object key -eq "MACHINE-M010.011").data -match `
    "operators=6-si-quan\+20-thuy-thu") "M010.011 sai thanh phan thuyen vien."
Assert-True (($machines | Where-Object key -eq "MACHINE-M010.020").data -match `
    "nonStateOperators=1-bac-5-10") "M011.020 thieu nhan cong ngoai ngan sach."
Assert-True (($machines | Where-Object key -eq "MACHINE-M010.024").data -match `
    "nonStateReferencePriceVnd=350000") "M011.024 sai nguyen gia ngoai ngan sach."
Assert-True (($machines | Where-Object key -eq "MACHINE-M010.033").pageTo -eq "29") `
    "Locator may chua bao phu Bang 03."

$costs = @(Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $sourceRoot "modules\CostRule.tsv"))
Assert-True (($costs | Where-Object key -eq "COST-K2-LINEAR").data -match `
    "ratesPercent=2.2,2.0,1.9,1.8,1.7") "K2 tuyen sai."
Assert-True (($costs | Where-Object key -eq "COST-K5-CIVIL").data -match `
    "thresholdBillion=10,20,50,100,200,500,1000,2000,5000,8000,10000") `
    "K5 chua cap nhat bang 2.24 TT38/2026."
Assert-True (($costs | Where-Object key -eq "COST-MIN-DISPOSAL-UNDER-4HA").data -match `
    "minimumVnd=6300000") "Chi phi huy toi thieu sai."

$geography = @(Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $sourceRoot "modules\Geography.tsv"))
$localities = @($geography | Where-Object key -like "GEO-LOCALITY-*")
Assert-Equal 34 $localities.Count "Bang dia ban khong du 34 tinh, thanh pho."
$provinces = @{}
foreach ($row in $localities) {
    $fields = @{}
    foreach ($field in $row.data.Split(";")) {
        $pair = $field.Split(@("="), 2, [StringSplitOptions]::None)
        $fields[$pair[0]] = $pair[1]
    }
    Assert-True (-not $provinces.ContainsKey($fields.province)) (
        "Trung dia ban " + $fields.province + ".")
    $provinces.Add($fields.province, $true)
    $zones = @($fields.zones.Split(","))
    Assert-Equal $zones.Count (@($zones | Select-Object -Unique).Count) (
        "Trung vung trong " + $fields.province + ".")
    Assert-True ($zones -contains $fields.fallbackZone) (
        "Fallback zone khong thuoc danh sach " + $fields.province + ".")
}

$audit = @(Import-Csv -Delimiter "`t" -Encoding UTF8 (
    Join-Path $packageRoot "audit\change-reasons.tsv"))
Assert-Equal 248 $audit.Count "Ma tran ly do thay doi khong phu het record diff."
foreach ($row in $audit) {
    Assert-True (-not [string]::IsNullOrWhiteSpace($row.reason)) ("Thieu ly do " + $row.key + ".")
    Assert-Equal "VerifiedAgainstOfficialSource" $row.reviewStatus (
        "Audit chua verified " + $row.key + ".")
}

$oldBundle = Join-Path $repoRoot "data\regulations\packages\BQP-RPBM-2021\bundle"
$diffOutput = @(& $tool diff $oldBundle $bundleRoot)
if ($LASTEXITCODE -ne 0) { throw "Record diff failed." }
Assert-True ($diffOutput -contains "RecordChanges=248;Added=50;Removed=0;Changed=198") `
    "Record diff summary khong khop."
$diffOutput | Set-Content -LiteralPath (Join-Path $packageRoot "audit\record-diff.txt") -Encoding UTF8

Write-Host "BQP 2025 package validation passed: 248 records; 6 official PDFs; 248 explained changes."
