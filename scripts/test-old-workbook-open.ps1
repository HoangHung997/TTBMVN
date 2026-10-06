param(
    [string]$WorkbookPath = ".\Dutoanmau\Checkpoints\DT-101-end.xlsm",
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$source = [IO.Path]::GetFullPath((Join-Path $repoRoot $WorkbookPath))
$variantRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau\Variants"))
$workingCopy = Join-Path $variantRoot "DT-605-old-workbook-working-copy.xlsm"
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

if (-not (Test-Path -LiteralPath $source)) { throw "Khong tim thay old workbook: $source" }
New-Item -ItemType Directory -Path $variantRoot -Force | Out-Null
if (Test-Path -LiteralPath $workingCopy) {
    if (-not $Overwrite) { throw "Working copy da ton tai: $workingCopy. Dung -Overwrite." }
    Remove-Item -LiteralPath $workingCopy -Force
}
Copy-Item -LiteralPath $source -Destination $workingCopy
$hashBefore = (Get-FileHash -Algorithm SHA256 -LiteralPath $workingCopy).Hash

[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$roleService = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$profileService = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$formType = $addin.GetType("ExcelAddIn1.Winform.FrmProjectSetup", $true)
$validateRoles = $roleService.GetMethod("Validate")
$tryLoadProfile = $profileService.GetMethod("TryLoad")

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$summary = $null
$form = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $roleArguments = New-Object "object[]" 1
    $roleArguments[0] = $workbook.PSObject.BaseObject
    $roleValidation = $validateRoles.Invoke($null, $roleArguments)
    if ($roleValidation.IsValid) { throw "Old workbook DT-101 khong duoc co san role marker." }

    $profileArguments = New-Object "object[]" 2
    $profileArguments[0] = $workbook.PSObject.BaseObject
    $profileArguments[1] = $null
    if ([bool]$tryLoadProfile.Invoke($null, $profileArguments)) {
        throw "Old workbook DT-101 khong duoc co san ProjectProfile."
    }

    $formArguments = New-Object "object[]" 2
    $formArguments[0] = $workbook.PSObject.BaseObject
    $formArguments[1] = $null
    $form = [Activator]::CreateInstance($formType, $formArguments)
    $saveField = $formType.GetField(
        "saveButton",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $mappingField = $formType.GetField(
        "mappingGrid",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $statusField = $formType.GetField(
        "statusLabel",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $saveButton = $saveField.GetValue($form)
    $mappingGrid = $mappingField.GetValue($form)
    $statusLabel = $statusField.GetValue($form)
    $mappedRows = @($mappingGrid.Rows | Where-Object {
        -not [string]::IsNullOrWhiteSpace([string]$_.Cells[1].Value)
    }).Count
    if (-not $saveButton.Enabled -or $mappingGrid.Rows.Count -ne 7 -or $mappedRows -ne 7 -or
        [string]::IsNullOrWhiteSpace([string]$statusLabel.Text)) {
        throw ("First-run form khong yeu cau anh xa old workbook dung cach. " +
            "SaveEnabled=$($saveButton.Enabled); MappingRows=$($mappingGrid.Rows.Count); Mapped=$mappedRows; " +
            "Status='$($statusLabel.Text)'.")
    }

    $summary = $workbook.Worksheets.Item("THKP-TC")
    $e27 = [double]$summary.Range("E27").Value2
    if ($e27 -ne 1198731000) { throw "Old workbook E27 regression: $e27" }
}
finally {
    if ($null -ne $form) { $form.Dispose() }
    Release-ComObject $summary
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}

$hashAfter = (Get-FileHash -Algorithm SHA256 -LiteralPath $workingCopy).Hash
if ($hashAfter -ne $hashBefore) { throw "Luong mo old workbook da ghi ngam vao file." }

[pscustomobject]@{
    Workbook = $workingCopy
    LegacyProfile = "ABSENT"
    LegacyRoles = "MISSING-AS-EXPECTED"
    FirstRunForm = "SUGGESTS-AND-REQUIRES-CONFIRMATION"
    E27 = 1198731000
    NoWrite = ($hashAfter -eq $hashBefore)
} | Format-List
Write-Host "DT-605 old workbook open PASS."
