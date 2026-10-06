param(
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-104-start.xlsm"
}
$testPath = [IO.Path]::GetFullPath($WorkbookPath)
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
$coreAssembly = [Reflection.Assembly]::LoadFrom($corePath)
$addinAssembly = [Reflection.Assembly]::LoadFrom($addinPath)

if (-not ("TTBMVN.Tests.RecordingSheetObserver" -as [type])) {
    Add-Type -ReferencedAssemblies @($corePath, $addinPath) -TypeDefinition @"
using System.Collections.Generic;
using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;

namespace TTBMVN.Tests
{
    public sealed class RecordingSheetObserver : IWorkbookSheetListObserver
    {
        public int CallbackCount { get; private set; }
        public IReadOnlyList<WorkbookSheetDescriptor> LastSheets { get; private set; }

        public void OnWorkbookSheetsChanged(IReadOnlyList<WorkbookSheetDescriptor> sheets)
        {
            CallbackCount++;
            LastSheets = sheets;
        }
    }
}
"@
}

$coordinatorType = $addinAssembly.GetType(
    "ExcelAddIn1.Funtion.WorkbookSheetChangeCoordinator",
    $true)
$formType = $addinAssembly.GetType("ExcelAddIn1.Winform.FrmDaodat", $true)
$descriptorType = $coreAssembly.GetType("ExcelAddIn1.Core.WorkbookSheetDescriptor", $true)
$sheetListType = $coreAssembly.GetType("ExcelAddIn1.Core.WorkbookSheetList", $true)
$observerType = "TTBMVN.Tests.RecordingSheetObserver" -as [type]

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Find-ByName([object]$Sheets, [string]$Name) {
    foreach ($sheet in $Sheets) {
        if ($sheet.Name -eq $Name) {
            return $sheet
        }
    }
    return $null
}

$excelSession = Start-IsolatedExcelTestProcess
$excel = $excelSession.Application
$workbook = $null
$testSheet = $null
$coordinator = $null
$subscription = $null
$oldDisplayAlerts = $excel.DisplayAlerts
try {
    $workbook = $excel.Workbooks.Open($testPath, 0, $false)
    $coordinator = [Activator]::CreateInstance($coordinatorType, @($excel.PSObject.BaseObject))
    $observer = [Activator]::CreateInstance($observerType)
    $subscribeMethod = $coordinatorType.GetMethod("Subscribe")
    $subscribeArgs = New-Object "object[]" 2
    $subscribeArgs[0] = $workbook
    $subscribeArgs[1] = $observer
    $subscription = $subscribeMethod.Invoke($coordinator, $subscribeArgs)
    $coordinator.PollNow()

    $initialCount = $observer.CallbackCount
    $initialSheetCount = $observer.LastSheets.Count
    if ($initialCount -lt 1) {
        throw "Coordinator khong gui snapshot ban dau."
    }

    $worksheets = $workbook.Worksheets
    $lastSheet = $null
    try {
        $lastSheet = $worksheets.Item($worksheets.Count)
        $testSheet = $worksheets.Add([Type]::Missing, $lastSheet, [Type]::Missing, [Type]::Missing)
    }
    finally {
        Release-ComObject $lastSheet
        Release-ComObject $worksheets
    }
    $testSheet.Name = "__DT104_EVENT_TEST"
    $coordinator.PollNow()
    if ($observer.LastSheets.Count -ne $initialSheetCount + 1 -or
        $null -eq (Find-ByName $observer.LastSheets "__DT104_EVENT_TEST")) {
        throw "Khong phat hien sheet moi."
    }

    $beforeRename = $observer.LastSheets
    $selectedBefore = Find-ByName $beforeRename "__DT104_EVENT_TEST"
    $testSheet.Name = "__DT104_RENAMED"
    $coordinator.PollNow()
    $afterRename = $observer.LastSheets
    $selectedAfter = Find-ByName $afterRename "__DT104_RENAMED"
    if ($null -eq $selectedAfter -or $selectedAfter.Key -ne $selectedBefore.Key) {
        throw "Khong phat hien rename theo stable key."
    }

    $resolveSelection = $sheetListType.GetMethod("ResolveSelection")
    $selectionArgs = New-Object "object[]" 3
    $selectionArgs[0] = $afterRename
    $selectionArgs[1] = $selectedBefore.Key
    $selectionArgs[2] = $selectedBefore.Name
    $resolved = $resolveSelection.Invoke($null, $selectionArgs)
    if ($resolved.Name -ne "__DT104_RENAMED") {
        throw "Selection khong bam theo stable key sau rename."
    }

    $uninitialized = [Runtime.Serialization.FormatterServices]::GetUninitializedObject($formType)
    $flags = [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
    $comboField = $formType.GetField("cboSheets", $flags)
    $tableField = $formType.GetField("txtTableData", $flags)
    $gridField = $formType.GetField("dgvTableData", $flags)
    $combo = New-Object Windows.Forms.ComboBox
    $tableText = New-Object Windows.Forms.TextBox
    $grid = New-Object Windows.Forms.DataGridView
    [void]$grid.Columns.Add("Data", "Data")
    [void]$grid.Rows.Add("UNSAVED_GRID_VALUE")
    $tableText.Text = "'SOURCE'!A1:D9"
    $comboField.SetValue($uninitialized, $combo)
    $tableField.SetValue($uninitialized, $tableText)
    $gridField.SetValue($uninitialized, $grid)
    foreach ($sheet in $beforeRename) {
        [void]$combo.Items.Add($sheet)
        if ($sheet.Key -eq $selectedBefore.Key) {
            $combo.SelectedItem = $sheet
        }
    }

    $applyMethod = $formType.GetMethod("ApplyWorkbookSheets", $flags)
    $applyArgs = New-Object "object[]" 2
    $applyArgs[0] = $afterRename
    $applyArgs[1] = $true
    [void]$applyMethod.Invoke($uninitialized, $applyArgs)
    if ($combo.SelectedItem.Name -ne "__DT104_RENAMED" -or
        $tableText.Text -ne "'SOURCE'!A1:D9" -or
        $grid.Rows[0].Cells[0].Value -ne "UNSAVED_GRID_VALUE") {
        throw "Refresh sheet list da lam thay doi state form khac."
    }

    $firstSheet = $workbook.Worksheets.Item(1)
    try {
        $firstSheet.Activate()
        $coordinator.PollNow()
    }
    finally {
        Release-ComObject $firstSheet
    }
    $callbackBeforeActivate = $observer.CallbackCount
    $testSheet.Activate()
    $coordinator.PollNow()
    if ($observer.CallbackCount -le $callbackBeforeActivate) {
        throw "SheetActivate khong danh dau refresh."
    }

    $excel.DisplayAlerts = $false
    $testSheet.Delete()
    Release-ComObject $testSheet
    $testSheet = $null
    $coordinator.PollNow()
    if ($observer.LastSheets.Count -ne $initialSheetCount -or
        $null -ne (Find-ByName $observer.LastSheets "__DT104_RENAMED")) {
        throw "Khong phat hien sheet bi xoa."
    }

    [pscustomobject]@{
        Workbook = $workbook.FullName
        InitialSheets = $initialSheetCount
        Callbacks = $observer.CallbackCount
        Add = "PASS"
        Rename = "PASS"
        Activate = "PASS"
        Delete = "PASS"
        SelectionByStableKey = "PASS"
        UnsavedRangeTextPreserved = "PASS"
        UnsavedGridPreserved = "PASS"
    } | Format-List
}
finally {
    $excel.DisplayAlerts = $false
    if ($null -ne $testSheet) {
        try { $testSheet.Delete() } catch { }
        Release-ComObject $testSheet
    }
    if ($null -ne $subscription) { $subscription.Dispose() }
    if ($null -ne $coordinator) { $coordinator.Dispose() }
    if ($null -ne $workbook) {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
    $excel.DisplayAlerts = $oldDisplayAlerts
    Stop-IsolatedExcelTestProcess $excelSession
}
