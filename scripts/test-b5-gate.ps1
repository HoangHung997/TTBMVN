param(
    [string]$CheckpointPath = "Dutoanmau\Checkpoints\DT-505-start.xlsm"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$checkpoint = [IO.Path]::GetFullPath((Join-Path $repoRoot $CheckpointPath))
if (-not (Test-Path -LiteralPath $checkpoint)) {
    throw "Khong tim thay checkpoint B5: $checkpoint"
}
$auditWorkingCopy = Join-Path $repoRoot "Dutoanmau\Variants\DT-505-audit-working-copy.xlsm"
New-Item -ItemType Directory -Path (Split-Path -Parent $auditWorkingCopy) -Force | Out-Null
Copy-Item -LiteralPath $checkpoint -Destination $auditWorkingCopy -Force

$tests = @(
    [pscustomobject]@{
        Name = "Release build + Core"
        Script = "build-release.ps1"
        Arguments = @()
    },
    [pscustomobject]@{
        Name = "DT-501 batch write"
        Script = "test-excel-batch-write.ps1"
        Arguments = @("-Overwrite")
    },
    [pscustomobject]@{
        Name = "DT-502 result audit"
        Script = "test-result-audit.ps1"
        Arguments = @("-WorkbookPath", $auditWorkingCopy)
    },
    [pscustomobject]@{
        Name = "DT-503 validation"
        Script = "test-workbook-validation.ps1"
        Arguments = @("-WorkbookPath", $checkpoint, "-Overwrite")
    },
    [pscustomobject]@{
        Name = "DT-504 migration transaction"
        Script = "test-package-migration-results.ps1"
        Arguments = @("-WorkbookPath", $checkpoint, "-Overwrite")
    },
    [pscustomobject]@{
        Name = "DT-504 migration UI"
        Script = "test-package-migration-ui.ps1"
        Arguments = @("-WorkbookPath", $checkpoint, "-Overwrite")
    },
    [pscustomobject]@{
        Name = "DT-505 runtime diagnostics"
        Script = "test-runtime-diagnostics.ps1"
        Arguments = @("-WorkbookPath", $checkpoint)
    },
    [pscustomobject]@{
        Name = "B5 semantic checkpoint"
        Script = "test-workbook-checkpoint.ps1"
        Arguments = @(
            "-CheckpointPath", $CheckpointPath,
            "-ExpectedE27", "1201557000",
            "-ExpectedFormulaErrors", "92")
    }
)

$results = New-Object System.Collections.Generic.List[object]
$gateWatch = [Diagnostics.Stopwatch]::StartNew()
foreach ($test in $tests) {
    $path = Join-Path $PSScriptRoot $test.Script
    Write-Host "`n=== $($test.Name) ===" -ForegroundColor Cyan
    $watch = [Diagnostics.Stopwatch]::StartNew()
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $path @($test.Arguments)
    $exitCode = $LASTEXITCODE
    $watch.Stop()
    if ($exitCode -ne 0) {
        throw "B5 gate failed at $($test.Name), exit=$exitCode."
    }
    $results.Add([pscustomobject]@{
        Test = $test.Name
        Status = "PASS"
        Seconds = [Math]::Round($watch.Elapsed.TotalSeconds, 1)
    })
}
$gateWatch.Stop()

Write-Host "`n=== B5 GATE PASSED ===" -ForegroundColor Green
$results | Format-Table -AutoSize
[pscustomobject]@{
    Checkpoint = $checkpoint
    Tests = $results.Count
    Status = "PASS"
    TotalSeconds = [Math]::Round($gateWatch.Elapsed.TotalSeconds, 1)
} | Format-List
