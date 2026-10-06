param(
    [string]$WorkbookPath = ".\Dutoanmau\Checkpoints\DT-604-start.xlsm"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$workbookPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $WorkbookPath))
$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
$screenshotPath = Join-Path $repoRoot "tmp\DT-604-estimate-support.png"
[void][Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$formType = $addin.GetType("ExcelAddIn1.Winform.FrmSupport", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)

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

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$workbook = $null
$form = $null
$shell = $null
$timer = $null
try {
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
    $arguments = New-Object "object[]" 1
    $arguments[0] = $workbook.PSObject.BaseObject
    $form = [Activator]::CreateInstance($formType, $arguments)
    $licenseStatus = Get-PrivateField $formType $form "lblLicenseStatusValue"
    $expiry = Get-PrivateField $formType $form "lblExpiryValue"
    if ([string]::IsNullOrWhiteSpace($licenseStatus.Text) -or
        [string]::IsNullOrWhiteSpace($expiry.Text)) {
        throw "FrmSupport khong bind activation status."
    }
    $licenseStatusText = $licenseStatus.Text
    $expiryText = $expiry.Text
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
    $bitmap = New-Object System.Drawing.Bitmap($form.ClientSize.Width, $form.ClientSize.Height)
    try {
        $form.DrawToBitmap(
            $bitmap,
            (New-Object System.Drawing.Rectangle(0, 0, $form.ClientSize.Width, $form.ClientSize.Height)))
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
    $form.Close()
    $form.Dispose()
    $form = $null

    $shellArguments = New-Object "object[]" 2
    $shellArguments[0] = $workbook.PSObject.BaseObject
    $shellArguments[1] = $null
    $shell = [Activator]::CreateInstance($shellType, $shellArguments)
    $supportButton = Get-PrivateField $shellType $shell "btnSupport"
    if ($supportButton.Name -ne "btnSupport") { throw "Shell thieu nut Ho tro." }

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 150
    $timer.Add_Tick({
        foreach ($openForm in [System.Windows.Forms.Application]::OpenForms) {
            if ($openForm.GetType().FullName -eq "ExcelAddIn1.Winform.FrmSupport") {
                $openForm.Close()
                $timer.Stop()
                break
            }
        }
    })
    $timer.Start()
    $showSupport = $shellType.GetMethod(
        "ShowSupport",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showSupport.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    if ($timer.Enabled) { throw "Shell support dialog khong mo/dong dung." }

    [pscustomobject]@{
        LicenseStatus = $licenseStatusText
        Expiry = $expiryText
        WorkbookBoundConstructor = "PASS"
        ShellButton = "PASS"
        DialogOpen = "PASS"
        Screenshot = $screenshotPath
    } | Format-List
}
finally {
    if ($null -ne $timer) { $timer.Stop(); $timer.Dispose() }
    if ($null -ne $form) { $form.Close(); $form.Dispose() }
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}

Write-Host "DT-604 estimate support UI PASS."
