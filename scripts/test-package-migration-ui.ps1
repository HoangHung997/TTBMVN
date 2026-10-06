param(
    [string]$WorkbookPath,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-504-start.xlsm"
}
$variantRoot = Join-Path $repoRoot "Dutoanmau\Variants"
$workingCopy = Join-Path $variantRoot "DT-504-ui-working-copy.xlsm"
$screenshotPath = Join-Path $repoRoot "tmp\DT-504-package-migration.png"
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"

New-Item -ItemType Directory -Path $variantRoot,(Split-Path -Parent $screenshotPath) -Force | Out-Null
foreach ($path in @($workingCopy, $screenshotPath)) {
    if (Test-Path -LiteralPath $path) {
        if (-not $Overwrite) { throw "File test da ton tai: $path. Dung -Overwrite." }
        Remove-Item -LiteralPath $path -Force
    }
}
Copy-Item -LiteralPath ([IO.Path]::GetFullPath($WorkbookPath)) -Destination $workingCopy

$core = [Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$bootstrapType = $addin.GetType("ExcelAddIn1.Funtion.RegulationPackageBootstrapService", $true)
$pinType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookRegulationPackageService", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.PackageMigrationControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)
$loadPackages = $bootstrapType.GetMethod("LoadAvailablePackages")
$pin = $pinType.GetMethod("PinAndSave")

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-Static([Reflection.MethodInfo]$Method, [object[]]$Arguments) {
    $invokeArguments = New-Object "object[]" $Arguments.Length
    for ($index = 0; $index -lt $Arguments.Length; $index++) {
        $invokeArguments[$index] = if ($null -eq $Arguments[$index]) { $null } else { $Arguments[$index].PSObject.BaseObject }
    }
    return $Method.Invoke($null, $invokeArguments)
}

function Get-PrivateField([type]$Type, [object]$Instance, [string]$Name) {
    $field = $Type.GetField($Name, [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    if ($null -eq $field) { throw "Khong tim thay field $Name." }
    return $field.GetValue($Instance)
}

$packages = @($loadPackages.Invoke($null, (New-Object "object[]" 0)))
$package2021 = $packages | Where-Object PackageId -eq "BQP-RPBM-2021" | Select-Object -First 1
$package2025 = $packages |
    Where-Object PackageId -eq "BQP-RPBM-2025" |
    Sort-Object { [version]$_.DataVersion } -Descending |
    Select-Object -First 1
if ($null -eq $package2021 -or $null -eq $package2025) { throw "Thieu package 2021/2025." }

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$control = $null
$hostForm = $null
$shell = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    Invoke-Static $pin @($workbook, $package2021) | Out-Null

    $arguments = New-Object "object[]" 1
    $arguments[0] = $workbook.PSObject.BaseObject
    $control = [Activator]::CreateInstance($controlType, $arguments)
    if (-not $control.SelectTarget($package2025.PackageId, $package2025.DataVersion)) {
        throw "UI khong chon duoc package 2025."
    }
    $preview = $control.PreviewSelected()
    if ($null -eq $preview -or -not $preview.Impact.CanApply -or
        $preview.Impact.ValueChanges.Count -ne 7) {
        throw "UI preview migration khong dat."
    }
    $valueGrid = Get-PrivateField $controlType $control "valueGrid"
    $packageGrid = Get-PrivateField $controlType $control "packageGrid"
    $applyButton = Get-PrivateField $controlType $control "applyButton"
    if ($valueGrid.Rows.Count -ne 7 -or
        $packageGrid.Rows.Count -ne $preview.Plan.Diff.Changes.Count -or
        -not $applyButton.Enabled) {
        throw "UI khong bind dung report hoac nut apply."
    }

    $hostForm = New-Object System.Windows.Forms.Form
    $hostForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $hostForm.ShowInTaskbar = $false
    $hostForm.ClientSize = New-Object System.Drawing.Size(830, 614)
    $control.Dock = [System.Windows.Forms.DockStyle]::Fill
    [void]$hostForm.Controls.Add($control)
    $hostForm.Show()
    [System.Windows.Forms.Application]::DoEvents()
    $bitmap = New-Object System.Drawing.Bitmap(830, 614)
    try {
        $control.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, 830, 614)))
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
    $hostForm.Controls.Remove($control)
    $hostForm.Close()
    $hostForm.Dispose()
    $hostForm = $null

    $shellArguments = New-Object "object[]" 2
    $shellArguments[0] = $workbook.PSObject.BaseObject
    $shellArguments[1] = $null
    $shell = [Activator]::CreateInstance($shellType, $shellArguments)
    $showMethod = $shellType.GetMethod(
        "ShowPackageMigration",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.PackageMigrationControl") {
        throw "Shell Du toan khong mo dung PackageMigrationControl."
    }

    [pscustomobject]@{
        Target = "$($package2025.PackageId)@$($package2025.DataVersion)"
        ValueRows = $valueGrid.Rows.Count
        PackageRows = $packageGrid.Rows.Count
        ApplyEnabled = $applyButton.Enabled
        Shell = "PASS"
        Screenshot = $screenshotPath
    } | Format-List
}
finally {
    if ($null -ne $hostForm) { $hostForm.Close(); $hostForm.Dispose() }
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $control) { $control.Dispose() }
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}
