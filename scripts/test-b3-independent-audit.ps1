param()

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$packageRoot = Join-Path $repoRoot "data\regulations\packages\BQP-RPBM-2025"
$sourceRoot = Join-Path $packageRoot "source"
$bundleRoot = Join-Path $packageRoot "bundle"
$reportPath = Join-Path $packageRoot "audit\DT-305-independent-audit.json"
$stagingRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "tmp\dt305-independent-bundle"))
$allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "tmp")).TrimEnd('\') + '\'
$tool = Join-Path $repoRoot "ExcelAddIn1.RegulationTool\bin\Release\ExcelAddIn1.RegulationTool.exe"
$auditor = Join-Path $repoRoot "tools\regulations\audit_bqp_2025.py"
$python = (Get-Command python -ErrorAction Stop).Source

if (-not $stagingRoot.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Staging path is outside repository tmp: $stagingRoot"
}
if (Test-Path -LiteralPath $stagingRoot) {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force
}

try {
    & $tool build $sourceRoot $stagingRoot
    if ($LASTEXITCODE -ne 0) { throw "Deterministic staging build failed." }

    $expectedFiles = @(Get-ChildItem -LiteralPath $bundleRoot -Recurse -File)
    $actualFiles = @(Get-ChildItem -LiteralPath $stagingRoot -Recurse -File)
    if ($expectedFiles.Count -ne $actualFiles.Count) {
        throw "Deterministic bundle file count mismatch."
    }
    foreach ($expected in $expectedFiles) {
        $relative = $expected.FullName.Substring($bundleRoot.Length).TrimStart('\')
        $actual = Join-Path $stagingRoot $relative
        if (-not (Test-Path -LiteralPath $actual)) {
            throw "Deterministic bundle is missing: $relative"
        }
        $expectedHash = (Get-FileHash -LiteralPath $expected.FullName -Algorithm SHA256).Hash
        $actualHash = (Get-FileHash -LiteralPath $actual -Algorithm SHA256).Hash
        if ($expectedHash -ne $actualHash) {
            throw "Deterministic bundle hash mismatch: $relative"
        }
    }

    & $python $auditor
    if ($LASTEXITCODE -ne 0) { throw "Independent source audit failed." }
    $report = Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($report.status -ne "PASS" -or $report.errors.Count -ne 0 -or
        $report.warnings.Count -ne 0 -or $report.recordCount -ne 248 -or
        $report.changeReasonCount -ne 248 -or $report.checks -lt 6500) {
        throw "Independent audit report did not satisfy the B3 gate."
    }

    [pscustomobject]@{
        Status = $report.status
        Checks = $report.checks
        Records = $report.recordCount
        ChangeReasons = $report.changeReasonCount
        OfficialPdfs = $report.sourcePdfCount
        PackageChecksum = $report.packageChecksum
        DeterministicFiles = $expectedFiles.Count
        Errors = $report.errors.Count
        Warnings = $report.warnings.Count
        Report = $reportPath
    } | Format-List
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force
    }
}
