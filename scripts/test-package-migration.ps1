param(
    [string]$WorkbookPath,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$arguments = @{}
if (-not [string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $arguments.WorkbookPath = $WorkbookPath
}
if ($Overwrite) {
    $arguments.Overwrite = $true
}
& (Join-Path $PSScriptRoot "test-package-migration-results.ps1") @arguments
