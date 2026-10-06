param(
    [string]$WorkbookPath,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-407-start.xlsm"
}
$sourcePath = [IO.Path]::GetFullPath($WorkbookPath)
$variantRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau\Variants"))
$workingCopy = Join-Path $variantRoot "DT-407-working-copy.xlsm"
$screenshotPath = Join-Path $repoRoot "tmp\DT-407-cost-summary.png"
if (-not $workingCopy.StartsWith($variantRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "Working copy nam ngoai Variants."
}
if (Test-Path -LiteralPath $workingCopy) {
    if (-not $Overwrite) { throw "Working copy da ton tai. Dung -Overwrite de tao lai." }
    Remove-Item -LiteralPath $workingCopy -Force
}
New-Item -ItemType Directory -Path $variantRoot -Force | Out-Null
Copy-Item -LiteralPath $sourcePath -Destination $workingCopy

$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
$core = [Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$moneyWordsType = $core.GetType("ExcelAddIn1.Core.VietnameseMoneyWords", $true)
$serviceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookCostSummaryService", $true)
$writerType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookCostSummaryWriter", $true)
$cultureType = $addin.GetType("ExcelAddIn1.Funtion.ExcelCulture", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.CostSummaryControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)
$loadMethod = $serviceType.GetMethod("Load")
$previewMethod = $serviceType.GetMethod("Preview")
$writeMethod = $writerType.GetMethod("Apply")
$getCultureMethod = $cultureType.GetMethod("GetNumberCulture")

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

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

function Get-PrivateField([Type]$Type, [object]$Instance, [string]$Name) {
    return $Type.GetField(
        $Name,
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
    ).GetValue($Instance)
}

function Get-ExcelErrorCount([object]$Workbook) {
    [long]$total = 0
    $worksheets = $Workbook.Worksheets
    try {
        for ($index = 1; $index -le $worksheets.Count; $index++) {
            $worksheet = $worksheets.Item($index)
            try {
                foreach ($cellType in @(-4123, 2)) {
                    $used = $null
                    $errors = $null
                    try {
                        $used = $worksheet.UsedRange
                        try {
                            $errors = $used.SpecialCells($cellType, 16)
                            $total += [long]$errors.CountLarge
                        }
                        catch [Runtime.InteropServices.COMException] {
                        }
                    }
                    finally {
                        Release-ComObject $errors
                        Release-ComObject $used
                    }
                }
            }
            finally { Release-ComObject $worksheet }
        }
    }
    finally { Release-ComObject $worksheets }
    return $total
}

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$control = $null
$uiHost = $null
$shell = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $initialErrors = Get-ExcelErrorCount $workbook
    $context = Invoke-Static $loadMethod @($workbook)
    $request = $context.DefaultRequest
    if ($context.PackageId -ne "BQP-RPBM-2025" -or $context.PackageVersion -ne "2.0.1") {
        throw "Package context sai: $($context.PackageId)@$($context.PackageVersion)."
    }
    if ([decimal]$request.MaterialVnd -ne [decimal]12540585.6195888 -or
        [decimal]$request.LaborVnd -ne [decimal]268414820.1 -or
        [decimal]$request.MachineVnd -ne [decimal]558860046.122220) {
        throw "Khong nhap dung tong VL/NC/M tu Gia DT TC."
    }
    if ([decimal]$request.AreaHa -ne [decimal]10.46 -or
        $request.Terrain -ne "urban-residential" -or
        $request.ProjectKind.ToString() -ne "Linear" -or
        $request.ConstructionKind.ToString() -ne "AgricultureAndEnvironment" -or
        [decimal]$request.DisposalWeightKg -ne [decimal]1001 -or
        [decimal]$request.PreTaxIncomeRatePercent -ne [decimal]5.5 -or
        [decimal]$request.VatRatePercent -ne [decimal]8) {
        throw "Khong nhap dung cau hinh legacy THKP-TC."
    }

    $preview = Invoke-Static $previewMethod @($context, $request)
    $result = $preview.Result
    $expected = @{
        Direct = 839815452L; Common = 107365928L; PreTaxIncome = 52094976L
        Z = 999276356L; K = 114386486L; Q = 1113662842L
        Vat = 87893896L; H = 1201556738L; Rounded = 1201557000L
    }
    if ($result.DirectCost.DirectVnd -ne $expected.Direct -or
        $result.DirectCost.CommonVnd -ne $expected.Common -or
        $result.PreTaxIncomeVnd -ne $expected.PreTaxIncome -or
        $result.ZVnd -ne $expected.Z -or
        $result.OtherCostTotalVnd -ne $expected.K -or
        $result.BeforeTaxVnd -ne $expected.Q -or
        $result.VatVnd -ne $expected.Vat -or
        $result.AfterTaxVnd -ne $expected.H -or
        $result.RoundedAfterTaxVnd -ne $expected.Rounded) {
        throw "Ket qua preview THKP-TC sai."
    }
    $k5 = $result.FindRequired("K5")
    if ([decimal]$k5.CalculatedRatePercent -ne [decimal]2.598 -or $k5.AppliedAmountVnd -ne 25961200L) {
        throw "K5 khong dung bang hien hanh hoac sai co so Z."
    }

    $culture = Invoke-Static $getCultureMethod @($excel)
    $controlArguments = New-Object "object[]" 2
    $controlArguments[0] = $workbook.PSObject.BaseObject
    $controlArguments[1] = $culture
    $control = [Activator]::CreateInstance($controlType, $controlArguments)
    if ($null -eq $control.CurrentPreview -or
        $control.CurrentPreview.Result.RoundedAfterTaxVnd -ne $expected.Rounded) {
        throw "CostSummaryControl preview sai."
    }
    $uiHost = New-Object System.Windows.Forms.Form
    $uiHost.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $uiHost.ShowInTaskbar = $false
    $uiHost.ClientSize = New-Object System.Drawing.Size(830, 614)
    $control.Dock = [System.Windows.Forms.DockStyle]::Fill
    [void]$uiHost.Controls.Add($control)
    $uiHost.Show()
    [System.Windows.Forms.Application]::DoEvents()
    $bitmap = New-Object System.Drawing.Bitmap(830, 614)
    try {
        $control.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, 830, 614)))
        New-Item -ItemType Directory -Path (Split-Path -Parent $screenshotPath) -Force | Out-Null
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
    $uiHost.Controls.Remove($control)
    $uiHost.Close(); $uiHost.Dispose(); $uiHost = $null

    $shellArguments = New-Object "object[]" 2
    $shellArguments[0] = $workbook.PSObject.BaseObject
    $shellArguments[1] = $null
    $shell = [Activator]::CreateInstance($shellType, $shellArguments)
    $showMethod = $shellType.GetMethod(
        "ShowCostSummary",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.CostSummaryControl") {
        throw "Shell Du toan khong mo CostSummaryControl."
    }
    $shell.Dispose(); $shell = $null
    $control.Dispose(); $control = $null

    $writeResult = Invoke-Static $writeMethod @($workbook, $preview)
    if ($writeResult.WrittenAddress -ne "D10:E27;A28" -or
        $writeResult.Calculation.RoundedAfterTaxVnd -ne $expected.Rounded) {
        throw "Write result THKP-TC sai."
    }
    $summary = $workbook.Worksheets.Item("THKP-TC")
    $estimate = $workbook.Worksheets.Item("Gia DT TC")
    try {
        $expectedWords = [string]$moneyWordsType.GetMethod("ToWords").Invoke(
            $null,
            [object[]]@([long]$expected.Rounded))
        $actualWords = [string]$summary.Range("A28").Value2
        if ([decimal]$summary.Range("D22").Value2 -ne [decimal]2.598 -or
            $summary.Range("E22").Formula -ne "=ROUND(D22%*E16,0)" -or
            [long]$summary.Range("E27").Value2 -ne $expected.Rounded -or
            -not $actualWords.EndsWith($expectedWords, [StringComparison]::Ordinal)) {
            throw "Cong thuc hoac gia tri THKP-TC sau batch write sai."
        }
        $oldD23 = [string]$estimate.Range("D23").Formula
        $estimate.Range("D23").Value2 = 134
        $excel.CalculateFull()
        $changedTotal = [long]$summary.Range("E27").Value2
        if ($changedTotal -eq $expected.Rounded) {
            throw "THKP-TC khong cap nhat khi dau vao Gia DT TC thay doi."
        }
        $estimate.Range("D23").Formula = $oldD23
        $excel.CalculateFull()
        if ([long]$summary.Range("E27").Value2 -ne $expected.Rounded) {
            throw "THKP-TC khong tro lai ket qua goc sau khoi phuc dau vao."
        }
    }
    finally {
        Release-ComObject $estimate
        Release-ComObject $summary
    }

    $workbook.Save()
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $summary = $workbook.Worksheets.Item("THKP-TC")
    try {
        if ([long]$summary.Range("E27").Value2 -ne $expected.Rounded -or
            [decimal]$summary.Range("D22").Value2 -ne [decimal]2.598) {
            throw "Ket qua THKP-TC mat sau save/reopen."
        }
    }
    finally { Release-ComObject $summary }
    $finalErrors = Get-ExcelErrorCount $workbook
    if ($finalErrors -ne $initialErrors -or $finalErrors -ne 92) {
        throw "So loi Excel thay doi: initial=$initialErrors final=$finalErrors."
    }

    [pscustomobject]@{
        Workbook = $workingCopy
        Package = "$($context.PackageId)@$($context.PackageVersion)"
        AreaHa = $request.AreaHa
        Terrain = $request.Terrain
        Direct = $result.DirectCost.DirectVnd
        Z = $result.ZVnd
        K5Rate = $k5.CalculatedRatePercent
        OtherCosts = $result.OtherCostTotalVnd
        Vat = $result.VatVnd
        RoundedTotal = $result.RoundedAfterTaxVnd
        BatchWrite = "PASS"
        InputRecalculation = "PASS"
        SaveReopen = "PASS"
        UI = "PASS"
        Screenshot = $screenshotPath
        ExcelErrors = $finalErrors
    } | Format-List
}
finally {
    if ($null -ne $uiHost) { $uiHost.Close(); $uiHost.Dispose() }
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $control) { $control.Dispose() }
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}
