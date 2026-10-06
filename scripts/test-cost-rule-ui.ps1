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
$workingCopy = Join-Path $variantRoot "DT-403-working-copy.xlsm"
$screenshotPath = Join-Path $repoRoot "tmp\DT-403-cost-rule.png"
if (-not $workingCopy.StartsWith(
    $variantRoot.TrimEnd('\') + '\',
    [StringComparison]::OrdinalIgnoreCase)) {
    throw "Working copy nam ngoai Variants."
}
if (Test-Path -LiteralPath $workingCopy) {
    if (-not $Overwrite) {
        throw "Working copy da ton tai. Dung -Overwrite de tao lai."
    }
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
$profileType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$costServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookCostRuleService", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.CostRuleControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)
$loadPackagesMethod = $bootstrapType.GetMethod("LoadAvailablePackages")
$pinMethod = $pinType.GetMethod("PinAndSave")
$loadProfileMethod = $profileType.GetMethod("LoadRequired")
$loadCatalogMethod = $costServiceType.GetMethod("LoadPinnedCatalog")

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

function Get-PrivateField([type]$Type, [object]$Instance, [string]$Name) {
    $field = $Type.GetField(
        $Name,
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    if ($null -eq $field) {
        throw "Khong tim thay field $($Type.FullName).$Name."
    }
    return $field.GetValue($Instance)
}

function New-CostControl([object]$Context) {
    $arguments = New-Object "object[]" 1
    $arguments[0] = $Context.PSObject.BaseObject
    return [Activator]::CreateInstance($controlType, $arguments)
}

function Get-ExcelErrorCount([object]$Workbook) {
    [long]$total = 0
    $worksheets = $Workbook.Worksheets
    try {
        for ($index = 1; $index -le $worksheets.Count; $index++) {
            $worksheet = $worksheets.Item($index)
            try {
                foreach ($cellType in @(-4123, 2)) {
                    $errors = $null
                    try {
                        $errors = $worksheet.Cells.SpecialCells($cellType, 16)
                        $total += [long]$errors.CountLarge
                    }
                    catch [Runtime.InteropServices.COMException] {
                    }
                    finally {
                        Release-ComObject $errors
                    }
                }
            }
            finally {
                Release-ComObject $worksheet
            }
        }
    }
    finally {
        Release-ComObject $worksheets
    }
    return $total
}

$packages = $loadPackagesMethod.Invoke($null, (New-Object "object[]" 0))
$package2021 = $packages | Where-Object { $_.PackageId -eq "BQP-RPBM-2021" } | Select-Object -First 1
$package2025 = $packages |
    Where-Object { $_.PackageId -eq "BQP-RPBM-2025" } |
    Sort-Object { [version]$_.DataVersion } -Descending |
    Select-Object -First 1
if ($null -eq $package2021 -or $null -eq $package2025) {
    throw "Khong co du package BQP-RPBM-2021/2025."
}

$excelSession = Start-IsolatedExcelTestProcess
$excel = $excelSession.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
$verifyBook = $null
$control = $null
$shell = $null
try {
    $initialErrors = Get-ExcelErrorCount $workbook

    Invoke-Static $pinMethod @($workbook, $package2021) | Out-Null
    $legacyContext = Invoke-Static $loadCatalogMethod @($workbook)
    $legacyControl = New-CostControl $legacyContext
    try {
        if ($null -eq $legacyControl.CurrentResult -or
            $legacyControl.CurrentResult.Components.Count -ne 6) {
            throw "Package 2021 khong tinh du K1-K6."
        }
        if ($legacyContext.PackageId -ne "BQP-RPBM-2021") {
            throw "CostRule service khong dung package da pin 2021."
        }
    }
    finally {
        $legacyControl.Dispose()
    }

    Invoke-Static $pinMethod @($workbook, $package2025) | Out-Null
    $context = Invoke-Static $loadCatalogMethod @($workbook)
    $control = New-CostControl $context
    $grid = Get-PrivateField $controlType $control "componentGrid"
    if ($null -eq $control.CurrentResult -or $grid.Rows.Count -ne 6) {
        throw "UI CostRule khong tinh du 6 thanh phan."
    }
    $originalK5 = [long]$control.CurrentResult.FindRequired("K5").CalculatedAmountVnd
    $k5Row = $grid.Rows | Where-Object { $_.Cells["colCode"].Value -eq "K5" } | Select-Object -First 1
    $k5Row.Cells["colOverride"].Value = $true
    $k5Row.Cells["colApplied"].Value = [long]12345678
    $k5Row.Cells["colReason"].Value = "Du toan rieng duoc phe duyet"
    $control.Recalculate()
    if ($null -eq $control.CurrentResult) {
        throw "UI CostRule tu choi override hop le."
    }
    $appliedK5 = $control.CurrentResult.FindRequired("K5")
    if (-not $appliedK5.IsOverridden -or
        [long]$appliedK5.AppliedAmountVnd -ne 12345678 -or
        [long]$appliedK5.CalculatedAmountVnd -ne $originalK5) {
        throw "UI CostRule khong giu dung calculated/applied override."
    }

    $control.Size = New-Object System.Drawing.Size(830, 614)
    $control.CreateControl()
    $bitmap = New-Object System.Drawing.Bitmap(830, 614)
    try {
        $control.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, 830, 614)))
        New-Item -ItemType Directory -Path (Split-Path -Parent $screenshotPath) -Force | Out-Null
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bitmap.Dispose()
    }

    $shellArguments = New-Object "object[]" 2
    $shellArguments[0] = $workbook.PSObject.BaseObject
    $shellArguments[1] = $null
    $shell = [Activator]::CreateInstance($shellType, $shellArguments)
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    $showCostMethod = $shellType.GetMethod(
        "ShowCostRule",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showCostMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.CostRuleControl") {
        throw "Shell Du toan khong mo CostRuleControl."
    }

    if ((Get-ExcelErrorCount $workbook) -ne $initialErrors) {
        throw "Tinh chi phi da lam thay doi so loi Excel."
    }
    $summary = $workbook.Worksheets.Item("THKP-TC")
    try {
        $summaryE27 = [double]$summary.Range("E27").Value2
    }
    finally {
        Release-ComObject $summary
    }
    if ($summaryE27 -ne 1198731000) {
        throw "Baseline E27 regression."
    }

    $shell.Dispose()
    $shell = $null
    $control.Dispose()
    $control = $null
    $workbook.Save()
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null
    $verifyBook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $profile = Invoke-Static $loadProfileMethod @($verifyBook)
    if ($profile.RegulationPackageId -ne "BQP-RPBM-2025" -or
        $profile.RegulationPackageVersion -ne "2.0.1") {
        throw "Pinned package khong ton tai sau save/reopen."
    }

    [pscustomobject]@{
        Workbook = $workingCopy
        Package = $profile.RegulationPackageId + "@" + $profile.RegulationPackageVersion
        LegacyPackageEngine = "PASS"
        Components = 6
        OverrideWithReason = "PASS"
        OriginalK5 = $originalK5
        AppliedK5 = [long]$appliedK5.AppliedAmountVnd
        ShellNavigation = "PASS"
        Screenshot = $screenshotPath
        SaveReopen = "PASS"
        ExcelErrors = $initialErrors
        SummaryE27 = $summaryE27
    } | Format-List
}
finally {
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $control) { $control.Dispose() }
    if ($null -ne $verifyBook) {
        try { $verifyBook.Close($false) } catch {}
        Release-ComObject $verifyBook
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    Stop-IsolatedExcelTestProcess $excelSession
}
