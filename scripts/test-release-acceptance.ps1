param(
    [string]$CheckpointPath = ".\Dutoanmau\Checkpoints\DT-605-start.xlsm",
    [switch]$Overwrite,
    [switch]$Resume
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$checkpoint = [IO.Path]::GetFullPath((Join-Path $repoRoot $CheckpointPath))
$evidencePath = Join-Path $repoRoot "tmp\DT-605-release-acceptance.json"
if (-not (Test-Path -LiteralPath $checkpoint)) { throw "Khong tim thay checkpoint: $checkpoint" }
$variantRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau\Variants"))
$auditWorkingCopy = Join-Path $variantRoot "DT-605-audit-working-copy.xlsm"
New-Item -ItemType Directory -Path $variantRoot -Force | Out-Null
if (Test-Path -LiteralPath $auditWorkingCopy) {
    if (-not $Overwrite) { throw "Audit working copy da ton tai. Dung -Overwrite." }
    Remove-Item -LiteralPath $auditWorkingCopy -Force
}
Copy-Item -LiteralPath $checkpoint -Destination $auditWorkingCopy

$tests = @(
    [pscustomobject]@{ Name = "Release build and Core"; Script = "build-release.ps1"; Arguments = @() },
    [pscustomobject]@{ Name = "Clean install upgrade reinstall"; Script = "test-release-lifecycle.ps1"; Arguments = @("-Overwrite") },
    [pscustomobject]@{ Name = "Open old workbook"; Script = "test-old-workbook-open.ps1"; Arguments = @("-Overwrite") },
    [pscustomobject]@{ Name = "Batch write transaction"; Script = "test-excel-batch-write.ps1"; Arguments = @("-Overwrite") },
    [pscustomobject]@{ Name = "Unit rate historical regression"; Script = "test-unit-rate.ps1"; Arguments = @("-WorkbookPath", (Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-405-start.xlsm"), "-Overwrite") },
    [pscustomobject]@{ Name = "Estimate appendix historical regression"; Script = "test-estimate-appendix.ps1"; Arguments = @("-WorkbookPath", (Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-406-start.xlsm"), "-Overwrite") },
    [pscustomobject]@{ Name = "Cost summary historical regression"; Script = "test-cost-summary.ps1"; Arguments = @("-WorkbookPath", (Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-407-start.xlsm"), "-Overwrite") },
    [pscustomobject]@{ Name = "Result audit regression"; Script = "test-result-audit.ps1"; Arguments = @("-WorkbookPath", $auditWorkingCopy) },
    [pscustomobject]@{ Name = "Workbook validation"; Script = "test-workbook-validation.ps1"; Arguments = @("-WorkbookPath", $checkpoint, "-Overwrite") },
    [pscustomobject]@{ Name = "Package migration"; Script = "test-package-migration-results.ps1"; Arguments = @("-WorkbookPath", $checkpoint, "-Overwrite") },
    [pscustomobject]@{ Name = "Update center UI"; Script = "test-update-center.ps1"; Arguments = @("-WorkbookPath", $CheckpointPath, "-Overwrite") },
    [pscustomobject]@{ Name = "Runtime diagnostics"; Script = "test-runtime-diagnostics.ps1"; Arguments = @("-WorkbookPath", $checkpoint) },
    [pscustomobject]@{ Name = "Estimate support UI"; Script = "test-estimate-support.ps1"; Arguments = @("-WorkbookPath", $CheckpointPath) },
    [pscustomobject]@{ Name = "Semantic checkpoint"; Script = "test-workbook-checkpoint.ps1"; Arguments = @("-CheckpointPath", $CheckpointPath, "-ExpectedE27", "1201557000", "-ExpectedFormulaErrors", "92") }
)

$results = New-Object System.Collections.Generic.List[object]
$checkpointHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $checkpoint).Hash
$startIndex = 0
$previousSeconds = 0.0
if ($Resume) {
    if (-not (Test-Path -LiteralPath $evidencePath)) {
        throw "Khong co evidence de resume: $evidencePath"
    }
    $previous = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
    if ($previous.CheckpointSha256 -ne $checkpointHash) {
        throw "Checkpoint da doi; khong duoc resume acceptance cu."
    }
    for ($index = 0; $index -lt [Math]::Min($tests.Count, $previous.Results.Count); $index++) {
        $old = $previous.Results[$index]
        if ($old.Test -ne $tests[$index].Name -or $old.Status -ne "PASS") { break }
        $results.Add([pscustomobject]@{
            Test = [string]$old.Test
            Status = "PASS"
            Seconds = [double]$old.Seconds
            Message = [string]$old.Message
        })
        $startIndex++
    }
    $previousSeconds = [double]($results | Measure-Object -Property Seconds -Sum).Sum
    Write-Host "Resume DT-605 after $startIndex passed tests." -ForegroundColor Yellow
}
$gateWatch = [Diagnostics.Stopwatch]::StartNew()
$failure = $null
for ($testIndex = $startIndex; $testIndex -lt $tests.Count; $testIndex++) {
    $test = $tests[$testIndex]
    $path = Join-Path $PSScriptRoot $test.Script
    Write-Host "`n=== $($test.Name) ===" -ForegroundColor Cyan
    $watch = [Diagnostics.Stopwatch]::StartNew()
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $path @($test.Arguments)
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) { throw "Exit code $exitCode." }
        $status = "PASS"
        $message = ""
    }
    catch {
        $status = "FAIL"
        $message = $_.Exception.Message
        $failure = "$($test.Name): $message"
    }
    finally {
        $watch.Stop()
        $results.Add([pscustomobject]@{
            Test = $test.Name
            Status = $status
            Seconds = [Math]::Round($watch.Elapsed.TotalSeconds, 1)
            Message = $message
        })
    }
    if ($null -ne $failure) { break }
}
$gateWatch.Stop()

$overall = if ($null -eq $failure -and $results.Count -eq $tests.Count) { "PASS" } else { "FAIL" }
$evidence = [ordered]@{
    SchemaVersion = 1
    Task = "DT-605"
    Status = $overall
    GeneratedAtUtc = [DateTime]::UtcNow.ToString("o")
    Checkpoint = $checkpoint
    CheckpointSha256 = $checkpointHash
    ExpectedTestCount = $tests.Count
    ExecutedTestCount = $results.Count
    TotalSeconds = [Math]::Round($previousSeconds + $gateWatch.Elapsed.TotalSeconds, 1)
    Results = $results.ToArray()
}
New-Item -ItemType Directory -Path (Split-Path -Parent $evidencePath) -Force | Out-Null
[IO.File]::WriteAllText(
    $evidencePath,
    ($evidence | ConvertTo-Json -Depth 8),
    (New-Object Text.UTF8Encoding($false)))

$results | Format-Table -AutoSize
Write-Host "Acceptance evidence: $evidencePath"
if ($overall -ne "PASS") { throw "DT-605 acceptance FAIL at $failure" }
Write-Host "DT-605 RELEASE ACCEPTANCE PASSED." -ForegroundColor Green
