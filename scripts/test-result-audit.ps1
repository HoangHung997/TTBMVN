param(
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Variants\DT-407-working-copy.xlsm"
}
$sourcePath = [IO.Path]::GetFullPath($WorkbookPath)
if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Khong tim thay workbook da ghi audit: $sourcePath"
}
$screenshotPath = Join-Path $repoRoot "tmp\DT-502-result-audit.png"

$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$auditType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookResultAuditService", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.ResultAuditControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)
$loadMethod = $auditType.GetMethod("Load")
$findMethod = $auditType.GetMethod("Find")
$findActiveMethod = $auditType.GetMethod("FindAtActiveCell")

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

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$estimateSheet = $null
$summarySheet = $null
$control = $null
$hostForm = $null
$shell = $null
try {
    $workbook = $excel.Workbooks.Open($sourcePath, 0, $false)
    $trail = Invoke-Static $loadMethod @($workbook)
    if ($trail.Entries.Count -ne 22) {
        throw "Audit phai co 22 entry (14 phu luc + 8 THKP), actual=$($trail.Entries.Count)."
    }

    $estimateSheet = $workbook.Worksheets.Item("Gia DT TC")
    $summarySheet = $workbook.Worksheets.Item("THKP-TC")
    $estimateCodeName = $estimateSheet.CodeName
    $summaryCodeName = $summarySheet.CodeName
    $lineAudit = Invoke-Static $findMethod @($workbook, $estimateCodeName, 22, 11)
    if ($null -eq $lineAudit -or
        $lineAudit.Kind.ToString() -ne "EstimateLine" -or
        $lineAudit.NormKey -ne "NORM-030.0300" -or
        $lineAudit.VariantCode -ne "water-3-12" -or
        $lineAudit.PackageId -ne "BQP-RPBM-2025" -or
        $lineAudit.PackageVersion -ne "2.0.1") {
        throw "Audit dong phu luc K22 sai."
    }
    $legal = @($lineAudit.Sources | Where-Object { $_.Kind.ToString() -eq "Regulation" })
    $prices = @($lineAudit.Sources | Where-Object { $_.Kind.ToString() -eq "Price" })
    if ($legal.Count -lt 1 -or $prices.Count -lt 1 -or
        [string]::IsNullOrWhiteSpace($legal[0].DocumentId) -or $legal[0].PageFrom -lt 1) {
        throw "Audit dong phu luc thieu locator van ban/trang hoac nguon gia."
    }

    $k5Audit = Invoke-Static $findMethod @($workbook, $summaryCodeName, 22, 5)
    if ($null -eq $k5Audit -or
        $k5Audit.Kind.ToString() -ne "CostComponent" -or
        $k5Audit.Label -ne "K5" -or
        -not ($k5Audit.Sources | Where-Object { $_.Kind.ToString() -eq "Regulation" })) {
        throw "Audit K5 khong truy dung can cu."
    }

    [void]$estimateSheet.Activate()
    [void]$estimateSheet.Range("K22").Select()
    $activeAudit = Invoke-Static $findActiveMethod @($workbook)
    if ($null -eq $activeAudit -or $activeAudit.EntryId -ne $lineAudit.EntryId) {
        throw "FindAtActiveCell khong tra dung dong phu luc."
    }

    $controlArguments = New-Object "object[]" 1
    $controlArguments[0] = $workbook.PSObject.BaseObject
    $control = [Activator]::CreateInstance($controlType, $controlArguments)
    if (-not $control.RefreshSelection() -or $control.CurrentEntry.EntryId -ne $lineAudit.EntryId) {
        throw "ResultAuditControl khong doc dung o dang chon."
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
        New-Item -ItemType Directory -Path (Split-Path -Parent $screenshotPath) -Force | Out-Null
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
    $hostForm.Controls.Remove($control)
    $hostForm.Close(); $hostForm.Dispose(); $hostForm = $null

    $shellArguments = New-Object "object[]" 2
    $shellArguments[0] = $workbook.PSObject.BaseObject
    $shellArguments[1] = $null
    $shell = [Activator]::CreateInstance($shellType, $shellArguments)
    $showMethod = $shellType.GetMethod(
        "ShowResultAudit",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.ResultAuditControl") {
        throw "Shell Du toan khong mo ResultAuditControl."
    }
    $shell.Dispose(); $shell = $null
    $control.Dispose(); $control = $null

    $estimateSheet.Name = "Gia DT Audit"
    $summarySheet.Name = "THKP Audit"
    $workbook.Save()
    Release-ComObject $summarySheet
    $summarySheet = $null
    Release-ComObject $estimateSheet
    $estimateSheet = $null
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $excel.Workbooks.Open($sourcePath, 0, $false)
    $reopenedTrail = Invoke-Static $loadMethod @($workbook)
    $reopenedLine = Invoke-Static $findMethod @($workbook, $estimateCodeName, 22, 11)
    $reopenedK5 = Invoke-Static $findMethod @($workbook, $summaryCodeName, 22, 5)
    if ($reopenedTrail.Entries.Count -ne 22 -or
        $reopenedLine.NormKey -ne "NORM-030.0300" -or
        $reopenedK5.Label -ne "K5") {
        throw "Audit mat hoac sai sau rename/save/reopen."
    }
    $estimateSheet = $workbook.Worksheets.Item("Gia DT Audit")
    $summarySheet = $workbook.Worksheets.Item("THKP Audit")
    $estimateSheet.Name = "Gia DT TC"
    $summarySheet.Name = "THKP-TC"
    $workbook.Save()
    Release-ComObject $summarySheet
    $summarySheet = $null
    Release-ComObject $estimateSheet
    $estimateSheet = $null
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $excel.Workbooks.Open($sourcePath, 0, $true)
    $finalTrail = Invoke-Static $loadMethod @($workbook)
    if ($finalTrail.Entries.Count -ne 22) {
        throw "Audit mat sau khi khoi phuc ten sheet."
    }

    [pscustomobject]@{
        Workbook = $sourcePath
        Entries = $trail.Entries.Count
        EstimateEntry = $lineAudit.EntryId
        Norm = "$($lineAudit.NormKey)/$($lineAudit.VariantCode)"
        LegalSource = "$($legal[0].DocumentId) trang $($legal[0].PageFrom)-$($legal[0].PageTo)"
        PriceSources = $prices.Count
        CostEntry = $k5Audit.EntryId
        ActiveCellLookup = "PASS"
        UI = "PASS"
        RenameSaveReopen = "PASS"
        Screenshot = $screenshotPath
    } | Format-List
}
finally {
    if ($null -ne $hostForm) { $hostForm.Close(); $hostForm.Dispose() }
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $control) { $control.Dispose() }
    Release-ComObject $summarySheet
    Release-ComObject $estimateSheet
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}
