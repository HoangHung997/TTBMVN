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
$workingCopy = Join-Path $variantRoot "DT-405-working-copy.xlsm"
$screenshotPath = Join-Path $repoRoot "tmp\DT-405-unit-rate.png"
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
$core = [Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$bootstrapType = $addin.GetType("ExcelAddIn1.Funtion.RegulationPackageBootstrapService", $true)
$pinType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookRegulationPackageService", $true)
$legacyType = $addin.GetType("ExcelAddIn1.Funtion.LegacyWorkbookPriceProfileService", $true)
$priceServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPriceProfileService", $true)
$unitServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookUnitRateService", $true)
$roleServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$excelCultureType = $addin.GetType("ExcelAddIn1.Funtion.ExcelCulture", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.UnitRateControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)
$bindingType = $core.GetType("ExcelAddIn1.Core.UnitRateResourceBinding", $true)
$validationExceptionType = $core.GetType("ExcelAddIn1.Core.UnitRateValidationException", $true)
$worksheetRoleType = $core.GetType("ExcelAddIn1.Core.WorksheetRole", $true)

$loadPackagesMethod = $bootstrapType.GetMethod("LoadAvailablePackages")
$pinMethod = $pinType.GetMethod("PinAndSave")
$importLegacyMethod = $legacyType.GetMethod("Import")
$savePriceMethod = $priceServiceType.GetMethod("SaveAndPin")
$loadContextMethod = $unitServiceType.GetMethod("LoadRequired")
$calculateMethod = $unitServiceType.GetMethod("Calculate")
$resolveSheetMethod = $roleServiceType.GetMethod("ResolveWorksheetRequired")
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

function Get-PrivateField([type]$Type, [object]$Instance, [string]$Name) {
    $field = $Type.GetField(
        $Name,
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    if ($null -eq $field) { throw "Khong tim thay field $($Type.FullName).$Name." }
    return $field.GetValue($Instance)
}

function Get-Cell([object]$Sheet, [string]$Address) {
    $cell = $Sheet.Range($Address)
    try { return $cell.Value2 }
    finally { Release-ComObject $cell }
}

function Assert-Decimal([decimal]$Expected, [decimal]$Actual, [decimal]$Tolerance, [string]$Message) {
    if ([Math]::Abs($Expected - $Actual) -gt $Tolerance) {
        throw "$Message Expected=$Expected Actual=$Actual"
    }
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

function New-BindingArray([string]$ResourceCode, [string]$PriceCode, [string]$Reason) {
    $result = [Array]::CreateInstance($bindingType, 1)
    $arguments = New-Object "object[]" 3
    $arguments[0] = $ResourceCode
    $arguments[1] = $PriceCode
    $arguments[2] = $Reason
    $result.SetValue([Activator]::CreateInstance($bindingType, $arguments), 0)
    Write-Output -NoEnumerate $result
}

function New-EmptyBindingArray {
    $result = [Array]::CreateInstance($bindingType, 0)
    Write-Output -NoEnumerate $result
}

function Invoke-UnitRate(
    [object]$Context,
    [string]$NormKey,
    [string]$Variant,
    [string[]]$Conditions,
    [Array]$Bindings) {
    $arguments = New-Object "object[]" 5
    $arguments[0] = $Context
    $arguments[1] = $NormKey
    $arguments[2] = $Variant
    $arguments[3] = $Conditions
    $arguments[4] = $Bindings
    return $calculateMethod.Invoke($null, $arguments)
}

function Get-RootException([Exception]$Exception) {
    $current = $Exception
    while ($null -ne $current.InnerException) { $current = $current.InnerException }
    return $current
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
$workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
$verifyBook = $null
$landSheet = $null
$waterSheet = $null
$unitControl = $null
$shell = $null
$uiHost = $null
try {
    $initialErrors = Get-ExcelErrorCount $workbook
    Invoke-Static $pinMethod @($workbook, $package2025) | Out-Null
    $priceProfile = Invoke-Static $importLegacyMethod @($workbook, "Ha Noi", $null)
    Invoke-Static $savePriceMethod @($workbook, $priceProfile) | Out-Null
    $context = Invoke-Static $loadContextMethod @($workbook)
    if ($context.NormCatalog.Definitions.Count -ne 34 -or $context.PriceProfile.Entries.Count -ne 34) {
        throw "WorkbookUnitRateContext khong nap dung catalog/profile."
    }

    $landRole = [Enum]::Parse($worksheetRoleType, "UnitRateLand")
    $waterRole = [Enum]::Parse($worksheetRoleType, "UnitRateWater")
    $landSheet = Invoke-Static $resolveSheetMethod @($workbook, $landRole)
    $waterSheet = Invoke-Static $resolveSheetMethod @($workbook, $waterRole)

    $emptyBindings = New-EmptyBindingArray
    $land = Invoke-UnitRate $context "NORM-020.0200" "density-1" ([string[]]@()) $emptyBindings
    Assert-Decimal ([decimal](Get-Cell $landSheet "F29")) $land.MaterialAmountVnd ([decimal]0.0001) "DG Can material"
    Assert-Decimal ([decimal](Get-Cell $landSheet "G29")) $land.LaborAmountVnd ([decimal]0.0001) "DG Can labor"
    Assert-Decimal ([decimal](Get-Cell $landSheet "H29")) $land.MachineAmountVnd ([decimal]0.0001) "DG Can machine"

    $water = Invoke-UnitRate $context "NORM-030.0100" "water-0.5-12" ([string[]]@()) $emptyBindings
    Assert-Decimal ([decimal](Get-Cell $waterSheet "F34")) $water.MaterialAmountVnd ([decimal]0.0001) "DG Nuoc material"
    Assert-Decimal ([decimal](Get-Cell $waterSheet "G34")) $water.LaborAmountVnd ([decimal]0.0001) "DG Nuoc labor"
    Assert-Decimal ([decimal](Get-Cell $waterSheet "H34")) $water.MachineAmountVnd ([decimal]0.0001) "DG Nuoc machine"

    $factored = Invoke-UnitRate $context "NORM-030.0100" "water-0.5-12" `
        ([string[]]@("current-gt-0-le-0.5")) $emptyBindings
    Assert-Decimal ([decimal]12969990) $factored.LaborAmountVnd ([decimal]0.0001) "He so labor"
    Assert-Decimal ([decimal]43531871.108) $factored.MachineAmountVnd ([decimal]0.0001) "He so machine"

    $bindingRequired = $false
    try {
        Invoke-UnitRate $context "NORM-030.0400" "water-0.5-12" ([string[]]@()) $emptyBindings | Out-Null
    }
    catch {
        $root = Get-RootException $_.Exception
        $bindingRequired = $validationExceptionType.IsInstanceOfType($root) -and
            ($root.Validation.Issues | Where-Object { $_.IssueCode.ToString() -eq "BindingRequired" }).Count -gt 0
    }
    if (-not $bindingRequired) { throw "M010.DIVING khong yeu cau binding ro rang." }

    $diving029 = New-BindingArray "M010.DIVING" "MACHINE-M010.029" "Bien phap lan 0,5m-3m"
    $diveShallow = Invoke-UnitRate $context "NORM-030.0400" "water-0.5-12" `
        ([string[]]@()) $diving029
    Assert-Decimal ([decimal](Get-Cell $waterSheet "F88")) $diveShallow.MaterialAmountVnd ([decimal]0.0001) "Dive material"
    Assert-Decimal ([decimal](Get-Cell $waterSheet "G88")) $diveShallow.LaborAmountVnd ([decimal]0.0001) "Dive labor"
    Assert-Decimal ([decimal](Get-Cell $waterSheet "H88")) $diveShallow.MachineAmountVnd ([decimal]0.0001) "Dive machine"

    $diving030 = New-BindingArray "M010.DIVING" "MACHINE-M010.030" "Bien phap lan 3m-6m"
    $diveDeep = Invoke-UnitRate $context "NORM-030.0600" "water-0.5-12" `
        ([string[]]@()) $diving030
    Assert-Decimal ([decimal](Get-Cell $waterSheet "F121")) $diveDeep.MaterialAmountVnd ([decimal]0.0001) "Deep material"
    Assert-Decimal ([decimal](Get-Cell $waterSheet "G121")) $diveDeep.LaborAmountVnd ([decimal]0.0001) "Deep labor"
    Assert-Decimal ([decimal](Get-Cell $waterSheet "H121")) $diveDeep.MachineAmountVnd ([decimal]0.0001) "Deep machine"

    $missingBlocked = $false
    try {
        Invoke-UnitRate $context "NORM-040.0500" "standard" ([string[]]@()) $emptyBindings | Out-Null
    }
    catch {
        $root = Get-RootException $_.Exception
        $missingBlocked = $validationExceptionType.IsInstanceOfType($root) -and
            ($root.Validation.Issues | Where-Object { $_.IssueCode.ToString() -eq "MissingPrice" }).Count -gt 0
    }
    if (-not $missingBlocked) { throw "Gia thieu khong bi chan." }

    $culture = Invoke-Static $getCultureMethod @($excel)
    $controlArguments = New-Object "object[]" 2
    $controlArguments[0] = $context
    $controlArguments[1] = $culture
    $unitControl = [Activator]::CreateInstance($controlType, $controlArguments)
    $unitControl.SetWaterMode($true)
    if (-not $unitControl.SelectDefinition("NORM-030.0400", "water-0.5-12") -or
        -not $unitControl.SetBinding("M010.DIVING", "MACHINE-M010.029", "Bien phap lan duoc duyet") -or
        -not $unitControl.SetCondition("current-gt-0-le-0.5", $true)) {
        throw "UnitRateControl khong chon duoc input test."
    }
    $uiResult = $unitControl.CalculateCurrent()
    if ($null -eq $uiResult -or $uiResult.Resources.Count -ne 4 -or
        $uiResult.MachineAmountVnd -ne [decimal]485809.654) {
        throw "UnitRateControl tinh sai ket qua."
    }

    $uiHost = New-Object System.Windows.Forms.Form
    $uiHost.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $uiHost.ShowInTaskbar = $false
    $uiHost.ClientSize = New-Object System.Drawing.Size(830, 614)
    $unitControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    [void]$uiHost.Controls.Add($unitControl)
    $uiHost.Show()
    [System.Windows.Forms.Application]::DoEvents()
    $bitmap = New-Object System.Drawing.Bitmap(830, 614)
    try {
        $unitControl.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, 830, 614)))
        New-Item -ItemType Directory -Path (Split-Path -Parent $screenshotPath) -Force | Out-Null
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
    $uiHost.Controls.Remove($unitControl)
    $uiHost.Close()
    $uiHost.Dispose()
    $uiHost = $null

    $shellArguments = New-Object "object[]" 2
    $shellArguments[0] = $workbook.PSObject.BaseObject
    $shellArguments[1] = $null
    $shell = [Activator]::CreateInstance($shellType, $shellArguments)
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    $showMethod = $shellType.GetMethod(
        "ShowUnitRate",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.UnitRateControl") {
        throw "Shell Du toan khong mo UnitRateControl."
    }

    $landSheet.Name = "DT405 Don Gia Can"
    $waterSheet.Name = "DT405 Don Gia Nuoc"
    $workbook.Save()
    $shell.Dispose(); $shell = $null
    $unitControl.Dispose(); $unitControl = $null
    Release-ComObject $landSheet; $landSheet = $null
    Release-ComObject $waterSheet; $waterSheet = $null
    $workbook.Close($false)
    Release-ComObject $workbook; $workbook = $null

    $verifyBook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $persistedContext = Invoke-Static $loadContextMethod @($verifyBook)
    if ($persistedContext.PriceProfile.Checksum -ne $priceProfile.Checksum -or
        $persistedContext.PackageId -ne "BQP-RPBM-2025") {
        throw "Unit rate context sai sau save/reopen."
    }
    $verifyLand = Invoke-Static $resolveSheetMethod @($verifyBook, $landRole)
    $verifyWater = Invoke-Static $resolveSheetMethod @($verifyBook, $waterRole)
    try {
        if ($verifyLand.Name -ne "DT405 Don Gia Can" -or
            $verifyWater.Name -ne "DT405 Don Gia Nuoc") {
            throw "Worksheet role khong bao toan sau rename/reopen."
        }
    }
    finally {
        Release-ComObject $verifyLand
        Release-ComObject $verifyWater
    }
    $summary = $verifyBook.Worksheets.Item("THKP-TC")
    try { $summaryE27 = [double](Get-Cell $summary "E27") }
    finally { Release-ComObject $summary }
    $finalErrors = Get-ExcelErrorCount $verifyBook
    if ($finalErrors -ne $initialErrors -or $summaryE27 -ne 1198731000) {
        throw "Workbook regression DT-405 khong dat."
    }

    [pscustomobject]@{
        Workbook = $workingCopy
        Context = $persistedContext.PackageId + " / " + $persistedContext.PriceProfile.ProfileId
        LandBaseline = "PASS"
        WaterBaseline = "PASS"
        CurrentFactor = "PASS"
        ExplicitMachineBinding = "PASS"
        MissingPriceBlocked = "PASS"
        RenamedRoleSheets = "PASS"
        SaveReopen = "PASS"
        UI = "PASS"
        Screenshot = $screenshotPath
        ExcelErrors = $finalErrors
        SummaryE27 = $summaryE27
    } | Format-List
}
finally {
    if ($null -ne $uiHost) { $uiHost.Close(); $uiHost.Dispose() }
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $unitControl) { $unitControl.Dispose() }
    Release-ComObject $landSheet
    Release-ComObject $waterSheet
    if ($null -ne $verifyBook) { $verifyBook.Close($false) }
    Release-ComObject $verifyBook
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}
