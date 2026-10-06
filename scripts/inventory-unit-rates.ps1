param(
    [string]$WorkbookPath,
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-405-start.xlsm"
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $repoRoot "tmp\DT-405-unit-rate-inventory.json"
}
$workbookPath = [IO.Path]::GetFullPath($WorkbookPath)
$outputPath = [IO.Path]::GetFullPath($OutputPath)
$allowedRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau"))
if (-not $workbookPath.StartsWith(
    $allowedRoot.TrimEnd('\') + '\',
    [StringComparison]::OrdinalIgnoreCase)) {
    throw "Workbook inventory nam ngoai Dutoanmau."
}

$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
$core = [Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$roleType = $core.GetType("ExcelAddIn1.Core.WorksheetRole", $true)
$roleServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$resolveMethod = $roleServiceType.GetMethod("ResolveWorksheetRequired")

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-Resolve([object]$Workbook, [string]$RoleName) {
    $role = [Enum]::Parse($roleType, $RoleName, $false)
    $arguments = New-Object "object[]" 2
    $arguments[0] = $Workbook.PSObject.BaseObject
    $arguments[1] = $role
    return $resolveMethod.Invoke($null, $arguments)
}

function Column-Name([int]$Column) {
    $value = $Column
    $name = ""
    while ($value -gt 0) {
        $value--
        $name = [char](65 + ($value % 26)) + $name
        $value = [Math]::Floor($value / 26)
    }
    return $name
}

function Array-Value([object]$Values, [int]$Row, [int]$Column) {
    if ($Values -is [Array]) {
        return $Values.GetValue($Row, $Column)
    }
    if ($Row -eq 1 -and $Column -eq 1) { return $Values }
    return $null
}

function Read-Sheet([object]$Workbook, [string]$RoleName) {
    $sheet = $null
    $used = $null
    try {
        $sheet = Invoke-Resolve $Workbook $RoleName
        $used = $sheet.UsedRange
        $values = $used.Value2
        $formulas = $used.Formula
        $formulasR1C1 = $used.FormulaR1C1
        $formats = $used.NumberFormat
        $startRow = [int]$used.Row
        $startColumn = [int]$used.Column
        $rowCount = [int]$used.Rows.Count
        $columnCount = [int]$used.Columns.Count
        $cells = New-Object System.Collections.Generic.List[object]
        for ($rowOffset = 1; $rowOffset -le $rowCount; $rowOffset++) {
            for ($columnOffset = 1; $columnOffset -le $columnCount; $columnOffset++) {
                $value = Array-Value $values $rowOffset $columnOffset
                $formula = Array-Value $formulas $rowOffset $columnOffset
                $formulaR1C1 = Array-Value $formulasR1C1 $rowOffset $columnOffset
                if ($null -eq $value -and $null -eq $formula) { continue }
                $absoluteRow = $startRow + $rowOffset - 1
                $absoluteColumn = $startColumn + $columnOffset - 1
                $cells.Add([ordered]@{
                    address = (Column-Name $absoluteColumn) + $absoluteRow
                    row = $absoluteRow
                    column = $absoluteColumn
                    value = $value
                    formula = if ($formula -is [string] -and $formula.StartsWith("=")) { $formula } else { $null }
                    formulaR1C1 = if ($formulaR1C1 -is [string] -and $formulaR1C1.StartsWith("=")) { $formulaR1C1 } else { $null }
                    numberFormat = [string](Array-Value $formats $rowOffset $columnOffset)
                })
            }
        }
        return [ordered]@{
            role = $RoleName
            name = [string]$sheet.Name
            codeName = [string]$sheet.CodeName
            usedAddress = [string]$used.Address($false, $false)
            startRow = $startRow
            startColumn = $startColumn
            rowCount = $rowCount
            columnCount = $columnCount
            cells = $cells
        }
    }
    finally {
        Release-ComObject $used
        Release-ComObject $sheet
    }
}

function Read-Names([object]$Workbook, [string[]]$SheetNames) {
    $result = New-Object System.Collections.Generic.List[object]
    $names = $null
    try {
        $names = $Workbook.Names
        for ($index = 1; $index -le $names.Count; $index++) {
            $name = $null
            try {
                $name = $names.Item($index)
                $refersTo = [string]$name.RefersTo
                $isRelevant = $false
                foreach ($sheetName in $SheetNames) {
                    if ($refersTo.IndexOf($sheetName, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                        $isRelevant = $true
                        break
                    }
                }
                if ($isRelevant) {
                    $result.Add([ordered]@{
                        name = [string]$name.Name
                        refersTo = $refersTo
                        visible = [bool]$name.Visible
                    })
                }
            }
            finally { Release-ComObject $name }
        }
    }
    finally { Release-ComObject $names }
    return $result
}

$excelSession = Start-IsolatedExcelTestProcess
$excel = $excelSession.Application
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
    $land = Read-Sheet $workbook "UnitRateLand"
    $water = Read-Sheet $workbook "UnitRateWater"
    $inventory = [ordered]@{
        workbook = $workbookPath
        generatedAtUtc = [DateTime]::UtcNow.ToString("o")
        sheets = @($land, $water)
        names = @(Read-Names $workbook @($land.name, $water.name))
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $outputPath) -Force | Out-Null
    $inventory | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $outputPath -Encoding UTF8
    [pscustomobject]@{
        Workbook = $workbookPath
        Land = "$($land.name) $($land.usedAddress) cells=$($land.cells.Count)"
        Water = "$($water.name) $($water.usedAddress) cells=$($water.cells.Count)"
        RelevantNames = $inventory.names.Count
        Output = $outputPath
    } | Format-List
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $excelSession
}
