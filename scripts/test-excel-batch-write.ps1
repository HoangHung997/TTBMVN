param(
    [string]$WorkbookPath,
    [switch]$Overwrite,
    [int]$Rows = 500,
    [int]$Columns = 8
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-501-start.xlsm"
}
$sourcePath = [IO.Path]::GetFullPath($WorkbookPath)
$variantRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau\Variants"))
$workingCopy = Join-Path $variantRoot "DT-501-working-copy.xlsm"
if (Test-Path -LiteralPath $workingCopy) {
    if (-not $Overwrite) { throw "Working copy da ton tai. Dung -Overwrite de tao lai." }
    Remove-Item -LiteralPath $workingCopy -Force
}
Copy-Item -LiteralPath $sourcePath -Destination $workingCopy

$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$contextType = $addin.GetType("ExcelAddIn1.Funtion.ExcelWriteContext", $true)
$transactionType = $addin.GetType("ExcelAddIn1.Funtion.ExcelBatchWriteTransaction", $true)
$writeValueMethod = $transactionType.GetMethod("WriteValue2")
$randomType = $addin.GetType("ExcelAddIn1.RandomDaodat", $true)
$writeDaodatMethod = $randomType.GetMethod(
    "WriteResultsToExcel",
    [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function New-WriteContext([object]$Application) {
    $arguments = New-Object "object[]" 2
    $arguments[0] = $Application.PSObject.BaseObject
    $arguments[1] = $false
    return [Activator]::CreateInstance($contextType, $arguments)
}

function Write-BatchValue([object]$Transaction, [object]$Range, [object]$Value) {
    $arguments = New-Object "object[]" 2
    $arguments[0] = $Range.PSObject.BaseObject
    $arguments[1] = $Value.PSObject.BaseObject
    $writeValueMethod.Invoke($Transaction, $arguments) | Out-Null
}

function Matrix([int]$RowCount, [int]$ColumnCount, [int]$Offset) {
    $matrix = New-Object 'object[,]' $RowCount,$ColumnCount
    for ($row = 0; $row -lt $RowCount; $row++) {
        for ($column = 0; $column -lt $ColumnCount; $column++) {
            $matrix[$row,$column] = $Offset + ($row * $ColumnCount) + $column + 1
        }
    }
    Write-Output -NoEnumerate $matrix
}

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$workbook = $null
$sheet = $null
$rollbackRange = $null
$batchRange = $null
$cellRange = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $sheet = $workbook.Worksheets.Add()
    $sheet.Name = "DT501_Batch_Test"

    $excel.ScreenUpdating = $true
    $excel.Calculation = -4105
    $excel.EnableEvents = $true
    $excel.DisplayAlerts = $true
    $stateRestored = $false
    try {
        $context = New-WriteContext $excel
        try {
            if ($excel.ScreenUpdating -or $excel.Calculation -ne -4135 -or
                $excel.EnableEvents -or $excel.DisplayAlerts) {
                throw "ExcelWriteContext khong dat dung trang thai ghi."
            }
            throw "Injected context failure"
        }
        finally { $context.Dispose() }
    }
    catch {
        if ($_.Exception.Message -ne "Injected context failure") { throw }
        $stateRestored = $excel.ScreenUpdating -and $excel.Calculation -eq -4105 -and
            $excel.EnableEvents -and $excel.DisplayAlerts
    }
    if (-not $stateRestored) { throw "ExcelWriteContext khong phuc hoi trang thai sau exception." }

    $rollbackRange = $sheet.Range("A1:C3")
    $seed = Matrix 3 3 100
    $rollbackRange.Value2 = $seed
    $before = $rollbackRange.Formula
    try {
        $transaction = [Activator]::CreateInstance($transactionType)
        try {
            Write-BatchValue $transaction $rollbackRange (Matrix 3 3 900)
            throw "Injected batch failure"
        }
        finally { $transaction.Dispose() }
    }
    catch {
        if ($_.Exception.Message -ne "Injected batch failure") { throw }
    }
    $after = $rollbackRange.Formula
    for ($row = 1; $row -le 3; $row++) {
        for ($column = 1; $column -le 3; $column++) {
            if ($before.GetValue($row,$column) -ne $after.GetValue($row,$column)) {
                throw "Batch rollback sai tai $row,$column."
            }
        }
    }

    $listType = $writeDaodatMethod.GetParameters()[1].ParameterType
    $tupleType = $listType.GetGenericArguments()[0]
    $results = [Activator]::CreateInstance($listType)
    foreach ($values5 in @(
        @(2.00, 1.50, 1.20, 0.80, 1.00),
        @(2.10, 1.60, 1.30, 0.90, 1.10),
        @(2.20, 1.70, 1.40, 1.00, 1.20))) {
        $tuple = [Activator]::CreateInstance($tupleType, [object[]]$values5)
        $listType.GetMethod("Add").Invoke($results, [object[]]@($tuple)) | Out-Null
    }
    $summaryCell = $sheet.Range("B5")
    try {
        $arguments = New-Object "object[]" 7
        $arguments[0] = $summaryCell.PSObject.BaseObject
        $arguments[1] = $results
        $arguments[2] = "General"
        $arguments[3] = "General"
        $arguments[4] = "General"
        $arguments[5] = $excel.PSObject.BaseObject
        $arguments[6] = $false
        $writeDaodatMethod.Invoke($null, $arguments) | Out-Null
        $sheet.Range("A5:B8").Calculate()
        if ([long]$sheet.Range("A5").Value2 -ne 3 -or
            [long]$sheet.Range("A6").Value2 -ne 1 -or
            [string]$sheet.Range("B6").FormulaR1C1 -notmatch "ROUND" -or
            [decimal]$sheet.Range("C6").Value2 -ne [decimal]2.00 -or
            [decimal]$sheet.Range("G8").Value2 -ne [decimal]1.20) {
            throw "Batch writer ho dao khong bao toan output."
        }
    }
    finally { Release-ComObject $summaryCell }

    $batchRange = $sheet.Range(
        $sheet.Cells(10, 1),
        $sheet.Cells(9 + $Rows, $Columns))
    $cellRange = $sheet.Range(
        $sheet.Cells(10, $Columns + 2),
        $sheet.Cells(9 + $Rows, ($Columns * 2) + 1))
    $values = Matrix $Rows $Columns 0

    $batchWatch = [Diagnostics.Stopwatch]::StartNew()
    $context = New-WriteContext $excel
    try {
        $transaction = [Activator]::CreateInstance($transactionType)
        try {
            Write-BatchValue $transaction $batchRange $values
            if ($transaction.BatchCount -ne 1 -or $transaction.CellCount -ne ($Rows * $Columns)) {
                throw "Thong ke batch transaction sai."
            }
            $transaction.Commit()
        }
        finally { $transaction.Dispose() }
    }
    finally { $context.Dispose() }
    $batchWatch.Stop()

    $cellWatch = [Diagnostics.Stopwatch]::StartNew()
    $context = New-WriteContext $excel
    try {
        for ($row = 1; $row -le $Rows; $row++) {
            for ($column = 1; $column -le $Columns; $column++) {
                $cell = $null
                try {
                    $cell = $cellRange.Cells($row, $column)
                    $cell.Value2 = $values[($row - 1), ($column - 1)]
                }
                finally { Release-ComObject $cell }
            }
        }
    }
    finally { $context.Dispose() }
    $cellWatch.Stop()

    if ([long]$batchRange.Cells($Rows, $Columns).Value2 -ne ($Rows * $Columns) -or
        [long]$cellRange.Cells($Rows, $Columns).Value2 -ne ($Rows * $Columns)) {
        throw "Gia tri batch/cell write khong khop."
    }
    if ($batchWatch.ElapsedMilliseconds -ge $cellWatch.ElapsedMilliseconds) {
        throw "Batch write khong nhanh hon cell write. Batch=$($batchWatch.ElapsedMilliseconds) Cell=$($cellWatch.ElapsedMilliseconds)"
    }


    $ratio = if ($batchWatch.Elapsed.TotalMilliseconds -le 0) {
        [double]::PositiveInfinity
    }
    else {
        $cellWatch.Elapsed.TotalMilliseconds / $batchWatch.Elapsed.TotalMilliseconds
    }
    $sheet.Delete()
    Release-ComObject $sheet
    $sheet = $null
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null

    [pscustomobject]@{
        Workbook = $workingCopy
        Cells = $Rows * $Columns
        BatchOperations = 1
        CellOperations = $Rows * $Columns
        BatchMilliseconds = $batchWatch.ElapsedMilliseconds
        CellMilliseconds = $cellWatch.ElapsedMilliseconds
        Speedup = [Math]::Round($ratio, 2)
        ExceptionRollback = "PASS"
        ExcelStateRestore = "PASS"
        ValueParity = "PASS"
        DaodatBatchPath = "PASS"
    } | Format-List
}
finally {
    Release-ComObject $cellRange
    Release-ComObject $batchRange
    Release-ComObject $rollbackRange
    if ($null -ne $sheet) {
        try { $sheet.Delete() } catch {}
    }
    Release-ComObject $sheet
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}
