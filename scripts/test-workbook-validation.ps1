param(
    [string]$WorkbookPath,
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
if ([string]::IsNullOrWhiteSpace($WorkbookPath)) {
    $WorkbookPath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-503-start.xlsm"
}
$sourcePath = [IO.Path]::GetFullPath($WorkbookPath)
$variantRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot "Dutoanmau\Variants"))
$workingCopy = Join-Path $variantRoot "DT-503-working-copy.xlsm"
$screenshotPath = Join-Path $repoRoot "tmp\DT-503-workbook-validation.png"
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
$validationType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookValidationService", $true)
$priceServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPriceProfileService", $true)
$roleServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.WorkbookValidationControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)
$priceType = $core.GetType("ExcelAddIn1.Core.PriceProfile", $true)
$priceEntryType = $core.GetType("ExcelAddIn1.Core.PriceProfileEntry", $true)
$priceOverrideType = $core.GetType("ExcelAddIn1.Core.PriceProfileOverride", $true)
$resourcePricesRole = [ExcelAddIn1.Core.WorksheetRole]::ResourcePrices

$scanMethod = $validationType.GetMethod("Scan")
$navigateMethod = $validationType.GetMethod("NavigateTo")
$loadPriceMethod = $priceServiceType.GetMethod("LoadRequired")
$savePriceMethod = $priceServiceType.GetMethod("Save")
$tryReadPriceMethod = $priceServiceType.GetMethod("TryReadPayload")
$restorePriceMethod = $priceServiceType.GetMethod("RestorePayload")
$resolveSheetMethod = $roleServiceType.GetMethod("ResolveRequired")
$clearRoleMethod = $roleServiceType.GetMethod("ClearRole")
$setRoleMethod = $roleServiceType.GetMethod("SetRole")
$createPriceMethod = $priceType.GetMethod("Create")

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

function Get-Issues([object]$Report, [string]$Code) {
    return @($Report.Issues | Where-Object { $_.Code.ToString() -eq $Code })
}

function New-ProfileWithoutEntry([object]$Profile, [string]$Code) {
    $keptEntries = @($Profile.Entries | Where-Object Code -ne $Code)
    if ($keptEntries.Count -eq $Profile.Entries.Count) {
        throw "Khong co price entry de bo: $Code"
    }
    $entryArray = [Array]::CreateInstance($priceEntryType, $keptEntries.Count)
    for ($index = 0; $index -lt $keptEntries.Count; $index++) {
        $entryArray.SetValue($keptEntries[$index].PSObject.BaseObject, $index)
    }
    $overrideArray = [Array]::CreateInstance($priceOverrideType, $Profile.Overrides.Count)
    for ($index = 0; $index -lt $Profile.Overrides.Count; $index++) {
        $overrideArray.SetValue($Profile.Overrides[$index].PSObject.BaseObject, $index)
    }
    $arguments = New-Object "object[]" 9
    $arguments[0] = $Profile.ProfileId
    $arguments[1] = $Profile.DataVersion
    $arguments[2] = $Profile.DisplayName
    $arguments[3] = $Profile.Location
    $arguments[4] = $Profile.ValuationDate
    $arguments[5] = $Profile.LaborAudience
    $arguments[6] = $Profile.CreatedAtUtc
    $arguments[7] = $entryArray
    $arguments[8] = $overrideArray
    return $createPriceMethod.Invoke($null, $arguments)
}

$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false
$workbook = $null
$estimateSheet = $null
$resourceSheet = $null
$brokenName = $null
$control = $null
$hostForm = $null
$shell = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $clean = Invoke-Static $scanMethod @($workbook)
    if (-not $clean.IsValid -or $clean.ErrorCount -ne 0) {
        $details = $clean.Issues | ForEach-Object { "$($_.Code):$($_.WorksheetRoleId)!$($_.Address):$($_.Message)" }
        throw "Checkpoint validation phai khong co loi chan. $($details -join ' | ')"
    }
    $baselineWarnings = $clean.WarningCount

    $profile = Invoke-Static $loadPriceMethod @($workbook)
    $payloadArgs = New-Object "object[]" 2
    $payloadArgs[0] = $workbook.PSObject.BaseObject
    $payloadArgs[1] = $null
    if (-not $tryReadPriceMethod.Invoke($null, $payloadArgs)) {
        throw "Khong doc duoc PriceProfile payload de test rollback."
    }
    $oldPricePayload = [string]$payloadArgs[1]
    $missingProfile = New-ProfileWithoutEntry $profile "MAT-CONCRETE-STAKE"
    Invoke-Static $savePriceMethod @($workbook, $missingProfile) | Out-Null
    $missingReport = Invoke-Static $scanMethod @($workbook)
    $missingIssues = Get-Issues $missingReport "MissingPrice"
    if ($missingIssues.Count -lt 1 -or -not ($missingIssues | Where-Object Address -eq "F13:N13")) {
        $details = $missingReport.Issues | ForEach-Object { "$($_.Code):$($_.Address):$($_.Message)" }
        throw "Khong bao dung MissingPrice tai dong 13. $($details -join ' | ')"
    }
    Invoke-Static $restorePriceMethod @($workbook, $oldPricePayload) | Out-Null

    $resourceSheet = Invoke-Static $resolveSheetMethod @($workbook, $resourcePricesRole)
    Invoke-Static $clearRoleMethod @($resourceSheet) | Out-Null
    $mappingReport = Invoke-Static $scanMethod @($workbook)
    $mappingIssues = Get-Issues $mappingReport "RoleMapping"
    if ($mappingIssues.Count -lt 1 -or -not ($mappingIssues | Where-Object WorksheetRoleId -eq "ResourcePrices")) {
        throw "Khong bao dung RoleMapping ResourcePrices."
    }
    Invoke-Static $setRoleMethod @($resourceSheet, $resourcePricesRole) | Out-Null
    Release-ComObject $resourceSheet
    $resourceSheet = $null

    $estimateSheet = $workbook.Worksheets.Item("Gia DT TC")
    $oldF13 = $estimateSheet.Range("F13").Value2
    $estimateSheet.Range("F13").Value2 = [double]$oldF13 + 1
    $excel.CalculateFull()
    $totalReport = Invoke-Static $scanMethod @($workbook)
    $totalIssues = Get-Issues $totalReport "IncorrectTotal"
    if ($totalIssues.Count -lt 1 -or -not ($totalIssues | Where-Object Address -eq "I27")) {
        throw "Khong bao dung IncorrectTotal I27 sau khi sua don gia."
    }
    $estimateSheet.Range("F13").Value2 = $oldF13
    $excel.CalculateFull()

    $oldK22Formula = [string]$estimateSheet.Range("K22").Formula
    $estimateSheet.Range("K22").Formula = "=D22*H21"
    $excel.CalculateFull()
    $formulaReport = Invoke-Static $scanMethod @($workbook)
    $formulaIssues = Get-Issues $formulaReport "StaleFormula"
    $k22Issue = $formulaIssues | Where-Object Address -eq "K22" | Select-Object -First 1
    if ($null -eq $k22Issue) {
        throw "Khong bao dung StaleFormula K22."
    }
    if (-not (Invoke-Static $navigateMethod @($workbook, $k22Issue)) -or
        $excel.ActiveCell.Address($false, $false) -ne "K22") {
        throw "Dieu huong issue khong chon dung K22."
    }
    $estimateSheet.Range("K22").Formula = $oldK22Formula
    $excel.CalculateFull()

    $brokenName = $workbook.Names.Add("TTBMVN_TEST_BROKEN", "=#REF!")
    $nameReport = Invoke-Static $scanMethod @($workbook)
    $nameIssues = Get-Issues $nameReport "BrokenName"
    if (-not ($nameIssues | Where-Object Subject -match "TTBMVN_TEST_BROKEN")) {
        throw "Khong bao defined name TTBMVN_TEST_BROKEN bi hong."
    }
    $brokenName.Delete()
    Release-ComObject $brokenName
    $brokenName = $null

    $final = Invoke-Static $scanMethod @($workbook)
    if (-not $final.IsValid -or $final.ErrorCount -ne 0 -or $final.WarningCount -ne $baselineWarnings) {
        throw "Workbook khong sach sau khi phuc hoi cac fault test."
    }

    $controlArguments = New-Object "object[]" 1
    $controlArguments[0] = $workbook.PSObject.BaseObject
    $control = [Activator]::CreateInstance($controlType, $controlArguments)
    $uiReport = $control.RunScan()
    if ($null -eq $uiReport -or -not $uiReport.IsValid) {
        throw "WorkbookValidationControl scan sai."
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
        "ShowWorkbookValidation",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.WorkbookValidationControl") {
        throw "Shell Du toan khong mo WorkbookValidationControl."
    }
    $shell.Dispose(); $shell = $null
    $control.Dispose(); $control = $null

    $workbook.Save()
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $reopened = Invoke-Static $scanMethod @($workbook)
    if (-not $reopened.IsValid -or $reopened.ErrorCount -ne 0 -or
        $reopened.WarningCount -ne $baselineWarnings) {
        throw "Validation thay doi sau save/reopen."
    }

    [pscustomobject]@{
        Workbook = $workingCopy
        CleanErrors = $clean.ErrorCount
        BaselineWarnings = $baselineWarnings
        MissingPrice = "PASS F13:N13"
        RoleMapping = "PASS ResourcePrices"
        IncorrectTotal = "PASS I27"
        StaleFormula = "PASS K22"
        BrokenName = "PASS TTBMVN_TEST_BROKEN"
        Navigation = "PASS K22"
        SaveReopen = "PASS"
        UI = "PASS"
        Screenshot = $screenshotPath
    } | Format-List
}
finally {
    if ($null -ne $brokenName) {
        try { $brokenName.Delete() } catch {}
    }
    Release-ComObject $brokenName
    if ($null -ne $hostForm) { $hostForm.Close(); $hostForm.Dispose() }
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $control) { $control.Dispose() }
    Release-ComObject $resourceSheet
    Release-ComObject $estimateSheet
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $workbook
    Stop-IsolatedExcelTestProcess $session
}
