param(
    [string]$WorkbookPath,
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-403-end.xlsm"
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repoRoot "tmp\DT-404-vl-nc-m-inventory.json"
}
$workbookPath = [IO.Path]::GetFullPath($WorkbookPath)
$allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau"))
if (-not $workbookPath.StartsWith(
    $allowedRoot.TrimEnd('\') + '\',
    [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook nam ngoai Dutoanmau: $workbookPath"
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Get-ArrayValue([object]$Value, [int]$Row, [int]$Column) {
    if ($Value -is [Array]) {
        return $Value.GetValue($Row, $Column)
    }
    if ($Row -eq 1 -and $Column -eq 1) {
        return $Value
    }
    return $null
}

function Get-ColumnName([int]$Column) {
    $name = ""
    while ($Column -gt 0) {
        $Column--
        $name = [char](65 + ($Column % 26)) + $name
        $Column = [Math]::Floor($Column / 26)
    }
    return $name
}

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$worksheet = $null
$usedRange = $null
try {
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
    $worksheet = $workbook.Worksheets.Item("VL-NC-M")
    $usedRange = $worksheet.UsedRange
    $rowCount = [int]$usedRange.Rows.Count
    $columnCount = [int]$usedRange.Columns.Count
    $firstRow = [int]$usedRange.Row
    $firstColumn = [int]$usedRange.Column
    $values = $usedRange.Value2
    $formulas = $usedRange.Formula
    $formats = $usedRange.NumberFormat
    $cells = New-Object System.Collections.Generic.List[object]
    for ($row = 1; $row -le $rowCount; $row++) {
        for ($column = 1; $column -le $columnCount; $column++) {
            $value = Get-ArrayValue $values $row $column
            $formula = Get-ArrayValue $formulas $row $column
            if ($null -eq $value -and $null -eq $formula) {
                continue
            }
            $absoluteRow = $firstRow + $row - 1
            $absoluteColumn = $firstColumn + $column - 1
            $cells.Add([pscustomobject]@{
                Address = (Get-ColumnName $absoluteColumn) + $absoluteRow
                Value = $value
                Formula = if ($formula -is [string] -and $formula.StartsWith("=")) { $formula } else { $null }
                NumberFormat = Get-ArrayValue $formats $row $column
            })
        }
    }

    $dependents = New-Object System.Collections.Generic.List[object]
    $worksheets = $workbook.Worksheets
    try {
        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $candidate = $worksheets.Item($sheetIndex)
            $candidateUsed = $null
            try {
                if ($candidate.Name -eq $worksheet.Name) {
                    continue
                }
                $candidateUsed = $candidate.UsedRange
                $candidateFormulas = $candidateUsed.Formula
                $candidateValues = $candidateUsed.Value2
                $candidateRows = [int]$candidateUsed.Rows.Count
                $candidateColumns = [int]$candidateUsed.Columns.Count
                for ($row = 1; $row -le $candidateRows; $row++) {
                    for ($column = 1; $column -le $candidateColumns; $column++) {
                        $formula = Get-ArrayValue $candidateFormulas $row $column
                        if (-not ($formula -is [string]) -or
                            $formula.IndexOf("VL-NC-M", [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                            continue
                        }
                        $dependents.Add([pscustomobject]@{
                            Sheet = $candidate.Name
                            Address = (Get-ColumnName ([int]$candidateUsed.Column + $column - 1)) +
                                ([int]$candidateUsed.Row + $row - 1)
                            LookupKey = if ([int]$candidateUsed.Column -le 2 -and
                                ([int]$candidateUsed.Column + $candidateColumns - 1) -ge 2) {
                                Get-ArrayValue $candidateValues $row (2 - [int]$candidateUsed.Column + 1)
                            }
                            else {
                                $null
                            }
                            Formula = $formula
                        })
                    }
                }
            }
            finally {
                Release-ComObject $candidateUsed
                Release-ComObject $candidate
            }
        }
    }
    finally {
        Release-ComObject $worksheets
    }

    $inventory = [pscustomobject]@{
        Workbook = $workbookPath
        Sheet = $worksheet.Name
        UsedRange = $usedRange.Address($false, $false)
        Rows = $rowCount
        Columns = $columnCount
        Cells = $cells
        ExternalDependents = $dependents
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
    $inventory | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
    [pscustomobject]@{
        Workbook = $workbookPath
        UsedRange = $inventory.UsedRange
        NonEmptyCells = $cells.Count
        ExternalDependents = $dependents.Count
        Output = [IO.Path]::GetFullPath($OutputPath)
    } | Format-List
}
finally {
    Release-ComObject $usedRange
    Release-ComObject $worksheet
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    Stop-IsolatedExcelTestProcess $session
}
