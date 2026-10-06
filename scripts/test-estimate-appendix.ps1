param(
    [string]$WorkbookPath,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\00. Du toan TP 3.xlsm"
}
$sourcePath = [IO.Path]::GetFullPath($WorkbookPath)
$variantRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau\Variants"))
$workingCopy = Join-Path $variantRoot "DT-406-working-copy.xlsm"
$screenshotPath = Join-Path $repoRoot "tmp\DT-406-estimate-appendix.png"
if (-not $workingCopy.StartsWith(
    $variantRoot.TrimEnd('\') + '\',
    [StringComparison]::OrdinalIgnoreCase)) {
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
[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$bootstrapType = $addin.GetType("ExcelAddIn1.Funtion.RegulationPackageBootstrapService", $true)
$pinType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookRegulationPackageService", $true)
$legacyType = $addin.GetType("ExcelAddIn1.Funtion.LegacyWorkbookPriceProfileService", $true)
$priceServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPriceProfileService", $true)
$unitServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookUnitRateService", $true)
$estimateServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookEstimateAppendixService", $true)
$estimateWriterType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookEstimateAppendixWriter", $true)
$auditServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookResultAuditService", $true)
$excelCultureType = $addin.GetType("ExcelAddIn1.Funtion.ExcelCulture", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.EstimateAppendixControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)

$loadPackagesMethod = $bootstrapType.GetMethod("LoadAvailablePackages")
$pinMethod = $pinType.GetMethod("PinAndSave")
$importLegacyMethod = $legacyType.GetMethod("Import")
$savePriceMethod = $priceServiceType.GetMethod("SaveAndPin")
$loadContextMethod = $unitServiceType.GetMethod("LoadRequired")
$importPlanMethod = $estimateServiceType.GetMethod("ImportLegacyPlan")
$previewMethod = $estimateServiceType.GetMethod("Preview")
$writeMethod = $estimateWriterType.GetMethod("Apply")
$loadAuditMethod = $auditServiceType.GetMethod("Load")
$findAuditMethod = $auditServiceType.GetMethod("Find")
$getCultureMethod = $excelCultureType.GetMethod("GetNumberCulture")

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

function Assert-Decimal([decimal]$Expected, [decimal]$Actual, [decimal]$Tolerance, [string]$Message) {
    if ([Math]::Abs($Expected - $Actual) -gt $Tolerance) {
        throw "$Message Expected=$Expected Actual=$Actual"
    }
}

function Invoke-Preview([object]$Context, [object]$Plan, [string[]]$Conditions) {
    $arguments = New-Object "object[]" 3
    $arguments[0] = $Context.PSObject.BaseObject
    $arguments[1] = $Plan.PSObject.BaseObject
    $arguments[2] = $Conditions
    return $previewMethod.Invoke($null, $arguments)
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

$packages = $loadPackagesMethod.Invoke($null, (New-Object "object[]" 0))
$package2025 = $packages |
    Where-Object { $_.PackageId -eq "BQP-RPBM-2025" } |
    Sort-Object { [version]$_.DataVersion } -Descending |
    Select-Object -First 1
if ($null -eq $package2025) { throw "Khong co package BQP-RPBM-2025." }

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$estimateControl = $null
$uiHost = $null
$shell = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $initialErrors = Get-ExcelErrorCount $workbook
    Invoke-Static $pinMethod @($workbook, $package2025) | Out-Null
    $priceProfile = Invoke-Static $importLegacyMethod @($workbook, "Ha Noi", $null)
    Invoke-Static $savePriceMethod @($workbook, $priceProfile) | Out-Null
    $context = Invoke-Static $loadContextMethod @($workbook)
    $plan = Invoke-Static $importPlanMethod @($workbook, $context)
    if ($plan.Lines.Count -ne 13) {
        throw "So dong phu luc imported sai: $($plan.Lines.Count)."
    }
    $expected = @(
        "12|NORM-010.0200|forest-1",
        "13|NORM-020.0200|density-1",
        "14|NORM-020.0300|soil-1",
        "15|NORM-020.0500|depth-3",
        "16|NORM-020.0500|depth-5",
        "17|NORM-020.0700|soil-2",
        "18|NORM-020.0800|",
        "20|NORM-030.0100|water-0.5-12",
        "21|NORM-030.0200|water-0.5-12",
        "22|NORM-030.0300|water-3-12",
        "23|NORM-030.0400|water-0.5-12",
        "24|NORM-030.0500|water-0.5-12",
        "25|NORM-030.0600|water-0.5-12"
    )
    $actual = @($plan.Lines | ForEach-Object { "$($_.TargetRow)|$($_.NormKey)|$($_.VariantCode)" })
    if (($actual -join ';') -ne ($expected -join ';')) {
        throw "Mapping phu luc sai. Actual=$($actual -join ';')"
    }
    $expectedBindingPrices = @{ 23 = "MACHINE-M010.029"; 24 = "MACHINE-M010.029"; 25 = "MACHINE-M010.030" }
    foreach ($row in @(23, 24, 25)) {
        $line = $plan.Lines | Where-Object TargetRow -eq $row | Select-Object -First 1
        if ($line.Bindings.Count -ne 1 -or
            $line.Bindings[0].ResourceCode -ne "M010.DIVING" -or
            $line.Bindings[0].PriceCode -ne $expectedBindingPrices[$row]) {
            throw "Khong import dung binding M010.DIVING tai row $row."
        }
    }

    $basePreview = Invoke-Preview $context $plan ([string[]]@())
    if (-not $basePreview.IsValid) {
        $messages = $basePreview.Lines | Where-Object IsBlocking | ForEach-Object { $_.Line.LineId + ': ' + $_.Error }
        throw "Preview co loi: $($messages -join ' | ')"
    }
    $currentPreview = Invoke-Preview $context $plan ([string[]]@("current-gt-0-le-0.5"))
    if (-not $currentPreview.IsValid) {
        $messages = $currentPreview.Lines | Where-Object IsBlocking | ForEach-Object { $_.Line.LineId + ': ' + $_.Error }
        throw "Preview he so co loi: $($messages -join ' | ')"
    }
    $targetSheet = $workbook.Worksheets.Item("Gia DT TC")
    try {
        $lineComparison = foreach ($resultLine in $currentPreview.Result.Lines) {
            $planLine = $plan.Lines | Where-Object LineId -eq $resultLine.Request.LineId | Select-Object -First 1
            $legacy = $targetSheet.Range("I$($planLine.TargetRow):K$($planLine.TargetRow)").Value2
            [pscustomobject]@{
                Row = $planLine.TargetRow
                Norm = "$($planLine.NormKey)/$($planLine.VariantCode)"
                CalcVL = $resultLine.MaterialAmountVnd
                LegacyVL = [decimal]$legacy.GetValue(1, 1)
                CalcNC = $resultLine.LaborAmountVnd
                LegacyNC = [decimal]$legacy.GetValue(1, 2)
                CalcM = $resultLine.MachineAmountVnd
                LegacyM = [decimal]$legacy.GetValue(1, 3)
            }
        }
        $lineComparison | Format-Table -AutoSize
    }
    finally { Release-ComObject $targetSheet }
    Assert-Decimal ([decimal]12540585.6195888) $currentPreview.Result.MaterialAmountVnd ([decimal]0.0001) "Tong VL"
    Assert-Decimal ([decimal]268414820.1) $currentPreview.Result.LaborAmountVnd ([decimal]0.0001) "Tong NC"
    Assert-Decimal ([decimal]558860046.122220) $currentPreview.Result.MachineAmountVnd ([decimal]0.0001) "Tong M"
    $waterMarking = $currentPreview.Lines | Where-Object { $_.Line.TargetRow -eq 22 }
    if ($waterMarking.IsBlocking -or $waterMarking.Error -notmatch "water-3-12") {
        throw "Dong 22 phai ghi nhan sua sai lech variant legacy 0,5-3m thanh 3-12m."
    }
    Assert-Decimal ([decimal]664425.135) `
        ($currentPreview.Result.MachineAmountVnd - [decimal]558195620.98722) `
        ([decimal]0.0001) "Chenh lech phap ly dong 22"
    $inactiveWarning = $currentPreview.Lines | Where-Object { $_.Line.TargetRow -eq 18 }
    if ($inactiveWarning.IsBlocking -or [string]::IsNullOrWhiteSpace($inactiveWarning.Error)) {
        throw "Dong zero quantity co gia thieu phai la canh bao khong blocking."
    }

    $culture = Invoke-Static $getCultureMethod @($excel)
    $controlArguments = New-Object "object[]" 2
    $controlArguments[0] = $workbook.PSObject.BaseObject
    $controlArguments[1] = $culture
    $estimateControl = [Activator]::CreateInstance($controlType, $controlArguments)
    if (-not $estimateControl.SetCondition("current-gt-0-le-0.5") -or
        $null -eq $estimateControl.CurrentPreview -or
        -not $estimateControl.CurrentPreview.IsValid -or
        $estimateControl.CurrentPreview.Result.MachineAmountVnd -ne [decimal]558860046.122220) {
        throw "EstimateAppendixControl preview sai."
    }
    $uiHost = New-Object System.Windows.Forms.Form
    $uiHost.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $uiHost.ShowInTaskbar = $false
    $uiHost.ClientSize = New-Object System.Drawing.Size(830, 614)
    $estimateControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    [void]$uiHost.Controls.Add($estimateControl)
    $uiHost.Show()
    [System.Windows.Forms.Application]::DoEvents()
    $bitmap = New-Object System.Drawing.Bitmap(830, 614)
    try {
        $estimateControl.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, 830, 614)))
        New-Item -ItemType Directory -Path (Split-Path -Parent $screenshotPath) -Force | Out-Null
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
    $uiHost.Controls.Remove($estimateControl)
    $uiHost.Close(); $uiHost.Dispose(); $uiHost = $null

    $shellArguments = New-Object "object[]" 2
    $shellArguments[0] = $workbook.PSObject.BaseObject
    $shellArguments[1] = $null
    $shell = [Activator]::CreateInstance($shellType, $shellArguments)
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    $showMethod = $shellType.GetMethod(
        "ShowEstimateAppendix",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.EstimateAppendixControl") {
        throw "Shell Du toan khong mo EstimateAppendixControl."
    }
    $shell.Dispose(); $shell = $null
    $estimateControl.Dispose(); $estimateControl = $null

    $writeResult = Invoke-Static $writeMethod @($workbook, $currentPreview)
    if ($writeResult.WrittenLineCount -ne 13 -or
        $writeResult.FirstRow -ne 11 -or
        $writeResult.LastRow -ne 27) {
        throw "Write result sai pham vi batch."
    }
    $audit = Invoke-Static $loadAuditMethod @($workbook)
    if ($audit.Entries.Count -ne 14) {
        throw "Audit phu luc phai co 13 dong va 1 block, actual=$($audit.Entries.Count)."
    }
    $targetSheet = $workbook.Worksheets.Item("Gia DT TC")
    try {
        $lineAudit = Invoke-Static $findAuditMethod @($workbook, $targetSheet.CodeName, 22, 11)
        if ($null -eq $lineAudit -or
            $lineAudit.NormKey -ne "NORM-030.0300" -or
            $lineAudit.VariantCode -ne "water-3-12" -or
            $lineAudit.PackageVersion -ne "2.0.1" -or
            $lineAudit.PriceProfileId -ne $priceProfile.ProfileId -or
            -not ($lineAudit.Sources | Where-Object Kind -eq ([ExcelAddIn1.Core.ResultAuditSourceKind]::Regulation))) {
            throw "Audit row 22 khong truy dung package/profile/norm/source."
        }
        if ($targetSheet.Range("I11").Formula -ne "=SUM(I12:I18)" -or
            $targetSheet.Range("I19").Formula -ne "=SUM(I20:I26)" -or
            $targetSheet.Range("K22").Formula -ne "=D22*H22" -or
            $targetSheet.Range("K26").Formula -ne "=SUM(K20:K25)*0.1" -or
            $targetSheet.Range("K27").Formula -ne "=K11+K19") {
            throw "Cong thuc phu luc sau batch write khong dung."
        }
        Assert-Decimal ([decimal]34270.8) ([decimal]$targetSheet.Range("H22").Value2) `
            ([decimal]0.0001) "Don gia may phap ly row 22"
        Assert-Decimal ([decimal]558860046.122220) ([decimal]$targetSheet.Range("K27").Value2) `
            ([decimal]0.0001) "Tong M sau ghi"

        $oldD23Formula = [string]$targetSheet.Range("D23").Formula
        $targetSheet.Range("D23").Value2 = 134
        $excel.CalculateFull()
        Assert-Decimal ([decimal]559383553.656220) ([decimal]$targetSheet.Range("K27").Value2) `
            ([decimal]0.0001) "Tong M sau doi dau vao"
        $targetSheet.Range("D23").Formula = $oldD23Formula
        $excel.CalculateFull()
        Assert-Decimal ([decimal]558860046.122220) ([decimal]$targetSheet.Range("K27").Value2) `
            ([decimal]0.0001) "Tong M sau khoi phuc dau vao"
    }
    finally { Release-ComObject $targetSheet }

    $workbook.Save()
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $reopenedContext = Invoke-Static $loadContextMethod @($workbook)
    $restoredPlan = Invoke-Static $importPlanMethod @($workbook, $reopenedContext)
    if ($restoredPlan.Lines.Count -ne 13 -or
        $restoredPlan.Lines[9].NormKey -ne "NORM-030.0300" -or
        $restoredPlan.Lines[9].VariantCode -ne "water-3-12") {
        throw "Khong khoi phuc dung plan phu luc tu audit sau reopen."
    }
    $verifySheet = $workbook.Worksheets.Item("Gia DT TC")
    try {
        Assert-Decimal ([decimal]558860046.122220) ([decimal]$verifySheet.Range("K27").Value2) `
            ([decimal]0.0001) "Tong M sau save/reopen"
        if ($verifySheet.Range("K26").Formula -ne "=SUM(K20:K25)*0.1") {
            throw "Cong thuc he so mat sau save/reopen."
        }
    }
    finally { Release-ComObject $verifySheet }
    $finalErrors = Get-ExcelErrorCount $workbook
    if ($finalErrors -ne $initialErrors -or $finalErrors -ne 92) {
        throw "So loi Excel thay doi: initial=$initialErrors final=$finalErrors."
    }

    [pscustomobject]@{
        Workbook = $workingCopy
        ImportedLines = $plan.Lines.Count
        AdjustmentRow = $plan.AdjustmentRow
        GrandTotalRow = $plan.GrandTotalRow
        LegacyBindings = ($plan.Lines | ForEach-Object { $_.Bindings.Count } | Measure-Object -Sum).Sum
        ZeroQuantityWarnings = ($currentPreview.Lines | Where-Object { -not $_.IsBlocking -and $_.Error }).Count
        MaterialTotal = $currentPreview.Result.MaterialAmountVnd
        LaborTotal = $currentPreview.Result.LaborAmountVnd
        MachineTotal = $currentPreview.Result.MachineAmountVnd
        Mapping = "PASS"
        Preview = "PASS"
        BatchWrite = "PASS"
        InputRecalculation = "PASS"
        SaveReopen = "PASS"
        Audit = "PASS"
        AuditPlanRestore = "PASS"
        UI = "PASS"
        Screenshot = $screenshotPath
        ExcelErrors = $finalErrors
    } | Format-List
}
finally {
    if ($null -ne $uiHost) { $uiHost.Close(); $uiHost.Dispose() }
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $estimateControl) { $estimateControl.Dispose() }
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}
