param(
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repoRoot "tmp\DT-501-excel-write-inventory.json"
}
$patterns = [ordered]@{
    Value2 = '\.Value2\s*='
    Formula = '\.Formula(?:R1C1)?\s*='
    Insert = '\.Insert\('
    Delete = '\.Delete\('
    WriteContext = 'new ExcelWriteContext'
    BatchTransaction = 'new ExcelBatchWriteTransaction'
}
$files = Get-ChildItem -LiteralPath (Join-Path $repoRoot "ExcelAddIn1") -Recurse -Filter "*.cs"
$matches = foreach ($file in $files) {
    $relative = $file.FullName.Substring($repoRoot.Length + 1)
    $lineNumber = 0
    foreach ($line in [IO.File]::ReadLines($file.FullName)) {
        $lineNumber++
        foreach ($entry in $patterns.GetEnumerator()) {
            if ($line -match $entry.Value) {
                [pscustomobject]@{
                    kind = $entry.Key
                    file = $relative
                    line = $lineNumber
                    code = $line.Trim()
                    cellIndexed = $line -match '\.Cells\s*\['
                }
            }
        }
    }
}
$summary = $matches | Group-Object kind | Sort-Object Name | ForEach-Object {
    [pscustomobject]@{ kind = $_.Name; count = $_.Count }
}
$report = [ordered]@{
    generatedAtUtc = [DateTime]::UtcNow.ToString("o")
    summary = @($summary)
    cellIndexedWrites = @($matches | Where-Object { $_.cellIndexed -and $_.kind -in @("Value2", "Formula") }).Count
    matches = @($matches)
}
$directory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
$report | ConvertTo-Json -Depth 4
