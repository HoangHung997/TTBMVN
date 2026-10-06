param(
    [string]$WorkbookPath = ".\Dutoanmau\Checkpoints\DT-602-start.xlsm",
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$sourceWorkbook = [IO.Path]::GetFullPath((Join-Path $repoRoot $WorkbookPath))
$variantRoot = Join-Path $repoRoot "Dutoanmau\Variants"
$workingCopy = Join-Path $variantRoot "DT-602-update-center-working-copy.xlsm"
$evidence = Join-Path $repoRoot "tmp\DT-601-BQP-RPBM-2025-2.0.1.ttbupdate"
$screenshotPath = Join-Path $repoRoot "tmp\DT-602-update-center.png"
$storeRoot = Join-Path $repoRoot "tmp\DT-602-package-store"
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"

New-Item -ItemType Directory -Path $variantRoot,(Split-Path -Parent $screenshotPath) -Force | Out-Null
foreach ($path in @($workingCopy, $screenshotPath)) {
    if (Test-Path -LiteralPath $path) {
        if (-not $Overwrite) { throw "File test da ton tai: $path. Dung -Overwrite." }
        Remove-Item -LiteralPath $path -Force
    }
}
if (Test-Path -LiteralPath $storeRoot) {
    if (-not $Overwrite) { throw "Store test da ton tai: $storeRoot. Dung -Overwrite." }
    $resolvedStore = [IO.Path]::GetFullPath($storeRoot)
    $resolvedTmp = [IO.Path]::GetFullPath((Join-Path $repoRoot "tmp"))
    if (-not $resolvedStore.StartsWith($resolvedTmp.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Store test nam ngoai tmp: $resolvedStore"
    }
    Remove-Item -LiteralPath $resolvedStore -Recurse -Force
}
if (-not (Test-Path -LiteralPath $sourceWorkbook)) { throw "Khong tim thay workbook: $sourceWorkbook" }
if (-not (Test-Path -LiteralPath $evidence)) { throw "Khong tim thay update evidence: $evidence" }
Copy-Item -LiteralPath $sourceWorkbook -Destination $workingCopy

[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$controlType = $addin.GetType("ExcelAddIn1.Winform.UpdateCenterControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)
$profileType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$loadProfile = $profileType.GetMethod("LoadRequired")
$bundle2021Path = Join-Path $repoRoot "data\regulations\packages\BQP-RPBM-2021\bundle"
$bundle2021 = [ExcelAddIn1.Core.RegulationPackageBundleReader]::Read($bundle2021Path)
$keys = [ExcelAddIn1.Core.OfflineUpdateTrustCatalog]::Production
$serviceArguments = New-Object "object[]" 3
$serviceArguments[0] = $storeRoot.PSObject.BaseObject
$serviceArguments[1] = $keys.PSObject.BaseObject
$serviceVersion = [Version]"1.0.0.0"
$serviceArguments[2] = $serviceVersion.PSObject.BaseObject
$keyEnumerableType = [System.Collections.Generic.IEnumerable[ExcelAddIn1.Core.OfflineUpdateTrustedKey]]
$serviceConstructor = [ExcelAddIn1.Core.OfflineUpdateCenterService].GetConstructor(
    [type[]]@([string], $keyEnumerableType, [Version]))
if ($null -eq $serviceConstructor) { throw "Khong tim thay OfflineUpdateCenterService constructor." }
$service = $serviceConstructor.Invoke($serviceArguments)

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Get-PrivateField([type]$Type, [object]$Instance, [string]$Name) {
    $field = $Type.GetField($Name, [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    if ($null -eq $field) { throw "Khong tim thay field $Name." }
    return $field.GetValue($Instance)
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
    $profileBefore = Invoke-Static $loadProfile @($workbook)

    $arguments = New-Object "object[]" 3
    $arguments[0] = $workbook.PSObject.BaseObject
    $arguments[1] = $service
    $arguments[2] = $bundle2021.Package
    $control = [Activator]::CreateInstance($controlType, $arguments)
    $inspection = $control.InspectFile($evidence)
    if ($null -eq $inspection -or
        $inspection.Verification.Disposition.ToString() -ne "Ready" -or
        $null -eq $inspection.Diff -or -not $inspection.Diff.HasChanges) {
        throw "Update center inspect/diff khong dat."
    }
    $diffGrid = Get-PrivateField $controlType $control "diffGrid"
    $installButton = Get-PrivateField $controlType $control "installButton"
    if ($diffGrid.Rows.Count -ne $inspection.Diff.Changes.Count -or -not $installButton.Enabled) {
        throw "UI khong bind dung diff hoac nut install."
    }

    $install = $control.InstallCurrent()
    if ($null -eq $install -or $install.Status.ToString() -ne "InstalledAndActivated") {
        throw "Update center install khong dat."
    }
    $profileAfter = Invoke-Static $loadProfile @($workbook)
    if ($profileAfter.RegulationPackageChecksum -ne $profileBefore.RegulationPackageChecksum) {
        throw "Install update da tu dong doi package workbook."
    }
    $restarted = $serviceConstructor.Invoke($serviceArguments)
    if ($restarted.ListInstalled().Count -ne 1 -or
        $restarted.ListPreferred().Count -ne 1 -or
        $restarted.ListPreferred()[0].PackageChecksum -ne $install.Package.PackageChecksum) {
        throw "Package/activation khong song qua restart service."
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
        "ShowUpdateCenter",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.UpdateCenterControl") {
        throw "Shell Du toan khong mo dung UpdateCenterControl."
    }

    [pscustomobject]@{
        Package = "$($install.Package.PackageId)@$($install.Package.DataVersion)"
        DiffRows = $diffGrid.Rows.Count
        Install = $install.Status.ToString()
        PreferredAfterRestart = $restarted.ListPreferred()[0].DataVersion
        WorkbookPinUnchanged = $true
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

Write-Host "DT-602 update center UI/install/restart PASS."
