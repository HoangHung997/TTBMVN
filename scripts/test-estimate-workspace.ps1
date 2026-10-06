param(
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$source = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-606-end.xlsm"
$variantRoot = Join-Path $repoRoot "Dutoanmau\Variants"
$workingCopy = Join-Path $variantRoot "DT-701-estimate-workspace.xlsm"
if (Test-Path -LiteralPath $workingCopy) {
    if (-not $Overwrite) { throw "Working copy da ton tai. Dung -Overwrite." }
    Remove-Item -LiteralPath $workingCopy -Force
}
Copy-Item -LiteralPath $source -Destination $workingCopy

[void][Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"))
$addin = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"))
$portfolioService = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPriceProfilePortfolioService", $true)
$workspaceService = $addin.GetType("ExcelAddIn1.Funtion.WorkbookEstimateWorkspaceService", $true)
$rangeReaderType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookEstimateRangeReader", $true)
$calculationService = $addin.GetType("ExcelAddIn1.Funtion.WorkbookEstimateCalculationService", $true)
$writerService = $addin.GetType("ExcelAddIn1.Funtion.WorkbookGeneratedEstimateWriter", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.EstimateAppendixControl", $true)
$rateControlType = $addin.GetType("ExcelAddIn1.Winform.EstimateRateCatalogControl", $true)
$cultureType = $addin.GetType("ExcelAddIn1.Funtion.ExcelCulture", $true)
$screenshotPath = Join-Path $repoRoot "tmp\DT-701-estimate-workspace.png"
$rateScreenshotPath = Join-Path $repoRoot "tmp\DT-701-rate-catalog.png"

function Invoke-Static([Reflection.MethodInfo]$Method, [object[]]$Arguments) {
    $invokeArguments = New-Object "object[]" $Arguments.Length
    for ($index = 0; $index -lt $Arguments.Length; $index++) {
        $invokeArguments[$index] = if ($null -eq $Arguments[$index]) {
            $null
        }
        else {
            $Arguments[$index].PSObject.BaseObject
        }
    }
    return $Method.Invoke($null, $invokeArguments)
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Set-Cell([object]$Sheet, [int]$Row, [int]$Column, [object]$Value) {
    $cell = $Sheet.Cells.Item($Row, $Column)
    try { $cell.Value2 = [string]$Value }
    finally { Release-ComObject $cell }
}

function New-Row(
    [string]$Id,
    [int]$Row,
    [string]$Code,
    [string]$Description,
    [decimal]$Quantity,
    [ExcelAddIn1.Core.EstimateWorkEnvironment]$Environment,
    [ExcelAddIn1.Core.MachineRateAudience]$Audience,
    [string]$Norm,
    [string]$Variant,
    [ExcelAddIn1.Core.UnitRateResourceBinding[]]$Bindings) {
    return [ExcelAddIn1.Core.EstimateWorkspaceRow]::new(
        $Id,
        "TTBMVN_PLDT_$Id",
        $Row,
        $Code,
        $Description,
        $(if ([string]::IsNullOrWhiteSpace($Norm)) { "" } else { "ha" }),
        $Quantity.ToString([Globalization.CultureInfo]::InvariantCulture),
        $Quantity,
        "",
        [decimal]0,
        $Environment,
        $Audience,
        $Norm,
        $Variant,
        [string[]]@(),
        $Bindings)
}

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$testSheet = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $portfolio = Invoke-Static $portfolioService.GetMethod("LoadRequired") @($workbook)
    $nonState = $portfolio.FindRequired([ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary)
    $state = [ExcelAddIn1.Core.PriceProfile]::Create(
        "$($nonState.ProfileId)-HLNS",
        $nonState.DataVersion,
        "$($nonState.DisplayName) - HLNS",
        $nonState.Location,
        $nonState.ValuationDate,
        [ExcelAddIn1.Core.MachineRateAudience]::StateBudgetSalary,
        [DateTime]::UtcNow,
        $nonState.Entries,
        $nonState.Overrides)
    $portfolio = [ExcelAddIn1.Core.PriceProfilePortfolio]::Create(
        [ExcelAddIn1.Core.PriceProfile[]]@($nonState, $state))
    Invoke-Static $portfolioService.GetMethod("Save") @($workbook, $portfolio) | Out-Null

    $legacySheet = $workbook.Worksheets.Item("Gia DT TC")
    $legacyRange = $legacySheet.Range("A9:L27")
    try {
        $legacyMap = [ExcelAddIn1.Core.EstimateColumnMap]::new(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
        $readerArguments = New-Object "object[]" 7
        $readerArguments[0] = $workbook.PSObject.BaseObject
        $readerArguments[1] = $legacyRange.PSObject.BaseObject
        $readerArguments[2] = $legacyMap
        $readerArguments[3] = $true
        $readerArguments[4] = [ExcelAddIn1.Core.EstimateWorkEnvironment]::Land
        $readerArguments[5] = [ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary
        $readerArguments[6] = $null
        $legacyWorkspace = $rangeReaderType.GetMethod("Read").Invoke($null, $readerArguments)
        if ($legacyWorkspace.Rows.Count -lt 5) {
            throw "Quet Gia DT TC!A9:L27 tra ve qua it dong."
        }
    }
    finally {
        Release-ComObject $legacyRange
        Release-ComObject $legacySheet
    }

    $testSheet = $workbook.Worksheets.Add()
    $testSheet.Name = "PLDT Test 701"
    $headers = @("Số hiệu", "Tên công tác", "ĐVT", "Khối lượng", "ĐG VL", "ĐG NC", "ĐG M", "TT VL", "TT NC", "TT M", "Ghi chú")
    for ($column = 1; $column -le $headers.Count; $column++) { Set-Cell $testSheet 1 $column $headers[$column - 1] }
    Set-Cell $testSheet 2 1 "CT.01"; Set-Cell $testSheet 2 2 "Dò tìm mật độ 1 - đoạn A"; Set-Cell $testSheet 2 3 "ha"; Set-Cell $testSheet 2 4 10
    Set-Cell $testSheet 3 1 "CT.02"; Set-Cell $testSheet 3 2 "Dò tìm mật độ 1 - đoạn B"; Set-Cell $testSheet 3 3 "ha"; Set-Cell $testSheet 3 4 20
    Set-Cell $testSheet 4 1 "CT.03"; Set-Cell $testSheet 4 2 "Dò tìm mật độ 2"; Set-Cell $testSheet 4 3 "ha"; Set-Cell $testSheet 4 4 30
    Set-Cell $testSheet 5 2 "KHU VỰC DƯỚI NƯỚC"; Set-Cell $testSheet 5 4 0
    Set-Cell $testSheet 6 1 "CT.04"; Set-Cell $testSheet 6 2 "Dò tìm mật độ 1 - HLNS"; Set-Cell $testSheet 6 3 "ha"; Set-Cell $testSheet 6 4 40
    Set-Cell $testSheet 7 1 "CT.05"; Set-Cell $testSheet 7 2 "Dò tìm dưới nước"; Set-Cell $testSheet 7 3 "signal"; Set-Cell $testSheet 7 4 50

    $errorCell = $testSheet.Range("K2")
    try {
        $errorCell.Formula = "=NA()"
        $errorCell.Calculate()
    }
    finally { Release-ComObject $errorCell }
    $readRange = $testSheet.Range("A1:K7")
    try {
        $errorMap = [ExcelAddIn1.Core.EstimateColumnMap]::new(1, 2, 3, 4, 11, 5, 6, 7, 8, 9, 10)
        $readerArguments = New-Object "object[]" 7
        $readerArguments[0] = $workbook.PSObject.BaseObject
        $readerArguments[1] = $readRange.PSObject.BaseObject
        $readerArguments[2] = $errorMap
        $readerArguments[3] = $true
        $readerArguments[4] = [ExcelAddIn1.Core.EstimateWorkEnvironment]::Land
        $readerArguments[5] = [ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary
        $readerArguments[6] = $null
        $errorWorkspace = $rangeReaderType.GetMethod("Read").Invoke($null, $readerArguments)
        if ($errorWorkspace.Rows[0].AcceptedQuantity -ne 0) {
            throw "Loi Excel o cot nghiem thu tuy chon khong duoc chuyen ve 0."
        }
    }
    finally { Release-ComObject $readRange }

    $columnMap = [ExcelAddIn1.Core.EstimateColumnMap]::new(1, 2, 3, 4, 0, 5, 6, 7, 8, 9, 10)
    $sourceBinding = [ExcelAddIn1.Core.EstimateSourceBinding]::new(
        "PLDT-TEST-701",
        $testSheet.Name,
        '$A$1:$K$7',
        2,
        7,
        $columnMap)
    $diving = [ExcelAddIn1.Core.UnitRateResourceBinding]::new(
        "M010.DIVING",
        "MACHINE-M010.029",
        "Thiet bi lan 0,5-3m theo bien phap duoc duyet")
    $rows = [ExcelAddIn1.Core.EstimateWorkspaceRow[]]@(
        (New-Row "r1" 2 "CT.01" "Do tim mat do 1 - doan A" 10 `
            ([ExcelAddIn1.Core.EstimateWorkEnvironment]::Land) `
            ([ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary) `
            "NORM-020.0200" "density-1" ([ExcelAddIn1.Core.UnitRateResourceBinding[]]@())),
        (New-Row "r2" 3 "CT.02" "Do tim mat do 1 - doan B" 20 `
            ([ExcelAddIn1.Core.EstimateWorkEnvironment]::Land) `
            ([ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary) `
            "NORM-020.0200" "density-1" ([ExcelAddIn1.Core.UnitRateResourceBinding[]]@())),
        (New-Row "r3" 4 "CT.03" "Do tim mat do 2" 30 `
            ([ExcelAddIn1.Core.EstimateWorkEnvironment]::Land) `
            ([ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary) `
            "NORM-020.0200" "density-2" ([ExcelAddIn1.Core.UnitRateResourceBinding[]]@())),
        (New-Row "r4" 5 "" "KHU VUC DUOI NUOC" 0 `
            ([ExcelAddIn1.Core.EstimateWorkEnvironment]::Water) `
            ([ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary) `
            "" "" ([ExcelAddIn1.Core.UnitRateResourceBinding[]]@())),
        (New-Row "r5" 6 "CT.04" "Do tim mat do 1 - HLNS" 40 `
            ([ExcelAddIn1.Core.EstimateWorkEnvironment]::Land) `
            ([ExcelAddIn1.Core.MachineRateAudience]::StateBudgetSalary) `
            "NORM-020.0200" "density-1" ([ExcelAddIn1.Core.UnitRateResourceBinding[]]@())),
        (New-Row "r6" 7 "CT.05" "Do tim duoi nuoc" 50 `
            ([ExcelAddIn1.Core.EstimateWorkEnvironment]::Water) `
            ([ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary) `
            "NORM-030.0400" "water-0.5-12" ([ExcelAddIn1.Core.UnitRateResourceBinding[]]@($diving)))
    )
    $workspace = [ExcelAddIn1.Core.EstimateWorkspace]::new($sourceBinding, $rows, [DateTime]::UtcNow)
    Invoke-Static $workspaceService.GetMethod("Save") @($workbook, $workspace) | Out-Null
    $preview = Invoke-Static $calculationService.GetMethod("Preview") @($workbook, $workspace)
    if (-not $preview.IsValid) {
        $issues = @($preview.Plan.Errors) + @($preview.Rates | Where-Object { -not $_.IsValid } | ForEach-Object Error)
        throw "Preview DT-701 loi: $($issues -join ' | ')"
    }
    if ($preview.Rates.Count -ne 4 -or $preview.CalculatedRowCount -ne 5 -or $preview.Plan.TextRows.Count -ne 1) {
        throw "Gom don gia sai: rates=$($preview.Rates.Count), rows=$($preview.CalculatedRowCount), text=$($preview.Plan.TextRows.Count)."
    }
    $shared = $preview.Rates | Where-Object {
        $_.Group.Identity.NormKey -eq "NORM-020.0200" -and
        $_.Group.Identity.VariantCode -eq "density-1" -and
        $_.Group.Identity.LaborAudience -eq [ExcelAddIn1.Core.MachineRateAudience]::NonStateSalary
    } | Select-Object -First 1
    if ($shared.Group.Rows.Count -ne 2) { throw "Hai cong tac cung density-1 khong dung chung don gia." }

    $written = Invoke-Static $writerService.GetMethod("Apply") @($workbook, $preview)
    if ($written.RateCount -ne 4 -or $written.LinkedRowCount -ne 5 -or $written.WorksheetNames.Count -ne 5) {
        throw "Writer sai: rates=$($written.RateCount), links=$($written.LinkedRowCount), sheets=$($written.WorksheetNames.Count)."
    }
    $rewritten = Invoke-Static $writerService.GetMethod("Apply") @($workbook, $preview)
    if ($rewritten.RateCount -ne 4 -or $rewritten.LinkedRowCount -ne 5 -or $rewritten.WorksheetNames.Count -ne 5) {
        throw "Writer lan hai khong idempotent: rates=$($rewritten.RateCount), links=$($rewritten.LinkedRowCount), sheets=$($rewritten.WorksheetNames.Count)."
    }
    $formulaF2 = [string]$testSheet.Range("E2").Formula
    $formulaF3 = [string]$testSheet.Range("E3").Formula
    $formulaF4 = [string]$testSheet.Range("E4").Formula
    if ($formulaF2 -ne $formulaF3 -or $formulaF2 -eq $formulaF4) {
        throw "Cong thuc link don gia dung chung/tach variant sai."
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$testSheet.Range("E5").Formula)) {
        throw "Dong van ban bi ghi don gia."
    }
    $testSheet.Name = "PLDT Test 701 Renamed"
    $workbook.Save()
    $workbook.Close($true)
    Release-ComObject $testSheet
    $testSheet = $null
    Release-ComObject $workbook
    $workbook = $null

    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $restored = Invoke-Static $workspaceService.GetMethod("LoadRequired") @($workbook)
    $reopenedPreview = Invoke-Static $calculationService.GetMethod("Preview") @($workbook, $restored)
    if (-not $reopenedPreview.IsValid -or $reopenedPreview.Rates.Count -ne 4) {
        throw "Workspace khong song qua save/reopen hoac rename sheet."
    }
    $culture = Invoke-Static $cultureType.GetMethod("GetNumberCulture") @($excel)
    $controlArguments = New-Object "object[]" 2
    $controlArguments[0] = $workbook.PSObject.BaseObject
    $controlArguments[1] = $culture
    $control = [Activator]::CreateInstance($controlType, $controlArguments)
    $hostForm = New-Object System.Windows.Forms.Form
    try {
        $hostForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $hostForm.ShowInTaskbar = $false
        $hostForm.ClientSize = New-Object System.Drawing.Size(1000, 660)
        $control.Dock = [System.Windows.Forms.DockStyle]::Fill
        [void]$hostForm.Controls.Add($control)
        $hostForm.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $populateMethod = $controlType.GetMethod(
            "PopulateGrid",
            [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
        [void]$populateMethod.Invoke($control, @($restored))
        [System.Windows.Forms.Application]::DoEvents()
        if ($null -eq $control.WorkspacePreview -or -not $control.WorkspacePreview.IsValid) {
            throw "EstimateAppendixControl khong nap dung workspace da luu."
        }
        $bitmap = New-Object System.Drawing.Bitmap(1000, 660)
        try {
            New-Item -ItemType Directory -Path (Split-Path -Parent $screenshotPath) -Force | Out-Null
            $control.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, 1000, 660)))
            $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $bitmap.Dispose() }
    }
    finally {
        $hostForm.Controls.Remove($control)
        $hostForm.Close(); $hostForm.Dispose()
        $control.Dispose()
    }
    $rateControl = [Activator]::CreateInstance($rateControlType, $controlArguments)
    $rateHostForm = New-Object System.Windows.Forms.Form
    try {
        $rateHostForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $rateHostForm.ShowInTaskbar = $false
        $rateHostForm.ClientSize = New-Object System.Drawing.Size(1000, 660)
        $rateControl.Dock = [System.Windows.Forms.DockStyle]::Fill
        [void]$rateHostForm.Controls.Add($rateControl)
        $rateHostForm.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $previewField = $rateControlType.GetField(
            "preview",
            [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
        $ratePreview = $previewField.GetValue($rateControl)
        if ($null -eq $ratePreview -or -not $ratePreview.IsValid -or $ratePreview.Rates.Count -ne 4) {
            throw "EstimateRateCatalogControl khong nap dung 4 don gia da gom."
        }
        $bitmap = New-Object System.Drawing.Bitmap(1000, 660)
        try {
            $rateControl.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, 1000, 660)))
            $bitmap.Save($rateScreenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $bitmap.Dispose() }
    }
    finally {
        $rateHostForm.Controls.Remove($rateControl)
        $rateHostForm.Close(); $rateHostForm.Dispose()
        $rateControl.Dispose()
    }
    $workbook.Close($false)
    Write-Host "DT-701 estimate workspace live test: PASS"
}
finally {
    if ($null -ne $testSheet) { Release-ComObject $testSheet }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    Stop-IsolatedExcelTestProcess $session
}
