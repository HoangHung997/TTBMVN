param(
    [string]$WorkbookPath = ".\Dutoanmau\Checkpoints\DT-606-start.xlsm"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$workbookPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $WorkbookPath))
[void][Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"))
$addin = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"))
$coordinatorType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookSheetChangeCoordinator", $true)
$formType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$workbook = $null
$worksheet = $null
$cell = $null
$coordinator = $null
$form = $null
try {
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
    $coordinatorArguments = New-Object "object[]" 1
    $coordinatorArguments[0] = $excel.PSObject.BaseObject
    $coordinator = [Activator]::CreateInstance($coordinatorType, $coordinatorArguments)

    $constructor = $formType.GetConstructor(
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic,
        $null,
        [Type[]]@(
            [Microsoft.Office.Interop.Excel.Workbook],
            $coordinatorType,
            [bool]),
        $null)
    if ($null -eq $constructor) { throw "Khong tim thay constructor test modeless." }
    $formArguments = New-Object "object[]" 3
    $formArguments[0] = $workbook.PSObject.BaseObject
    $formArguments[1] = $coordinator
    $formArguments[2] = $false
    $form = $constructor.Invoke($formArguments)
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
    if (-not $form.Visible -or $form.Modal) {
        throw "Form Du toan khong o trang thai modeless."
    }

    $worksheet = $workbook.Worksheets.Item(1)
    $cell = $worksheet.Range("B2")
    $worksheet.Activate()
    $cell.Select()
    [System.Windows.Forms.Application]::DoEvents()
    $activeAddress = [string]$excel.ActiveCell.Address($false, $false)
    if ($activeAddress -ne "B2" -or -not $form.Visible) {
        throw "Khong tuong tac duoc Excel khi form Du toan dang hien."
    }
    Release-ComObject $cell
    $cell = $null
    Release-ComObject $worksheet
    $worksheet = $null

    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null
    [System.Windows.Forms.Application]::DoEvents()
    if (-not $form.IsDisposed) {
        throw "Form Du toan khong tu dong dong khi workbook dong."
    }

    [pscustomobject]@{
        Modeless = "PASS"
        ExcelActiveCellWhileVisible = $activeAddress
        CloseWithWorkbook = "PASS"
    } | Format-List
}
finally {
    if ($null -ne $cell) { Release-ComObject $cell }
    if ($null -ne $worksheet) { Release-ComObject $worksheet }
    if ($null -ne $form -and -not $form.IsDisposed) { $form.Close(); $form.Dispose() }
    if ($null -ne $coordinator) { $coordinator.Dispose() }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}

Write-Host "Dutoan modeless live test: PASS"
