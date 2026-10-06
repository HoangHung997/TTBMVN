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
$workingCopy = Join-Path $variantRoot "DT-404-working-copy.xlsm"
$screenshotPath = Join-Path $repoRoot "tmp\DT-404-price-profile.png"
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
$profileType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$legacyType = $addin.GetType("ExcelAddIn1.Funtion.LegacyWorkbookPriceProfileService", $true)
$workbookPriceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPriceProfileService", $true)
$requirementType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookPriceRequirementService", $true)
$roleServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$controlType = $addin.GetType("ExcelAddIn1.Winform.PriceProfileControl", $true)
$shellType = $addin.GetType("ExcelAddIn1.Winform.Dutoan", $true)
$priceProfileType = $core.GetType("ExcelAddIn1.Core.PriceProfile", $true)
$priceOverrideType = $core.GetType("ExcelAddIn1.Core.PriceProfileOverride", $true)
$priceRequirementType = $core.GetType("ExcelAddIn1.Core.PriceRequirement", $true)
$coverageType = $core.GetType("ExcelAddIn1.Core.PriceProfileCoverageValidator", $true)
$worksheetRoleType = $core.GetType("ExcelAddIn1.Core.WorksheetRole", $true)

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

function Get-Cell([object]$Sheet, [string]$Address, [string]$Property) {
    $cell = $Sheet.Range($Address)
    try { return $cell.$Property }
    finally { Release-ComObject $cell }
}

function Find-RowByName([object]$Sheet, [string]$Name) {
    $range = $Sheet.Range("B1", "B500")
    try {
        $values = $range.Value2
        for ($row = 1; $row -le 500; $row++) {
            if ([string]::Equals(
                [string]$values.GetValue($row, 1),
                $Name,
                [StringComparison]::Ordinal)) {
                return $row
            }
        }
    }
    finally { Release-ComObject $range }
    throw "Khong tim thay dong $Name trong $($Sheet.Name)."
}

$loadPackagesMethod = $bootstrapType.GetMethod("LoadAvailablePackages")
$pinMethod = $pinType.GetMethod("PinAndSave")
$loadProjectMethod = $profileType.GetMethod("LoadRequired")
$importLegacyMethod = $legacyType.GetMethod("Import")
$applyLegacyMethod = $legacyType.GetMethod("ApplyDirectMaterialPrices")
$savePinMethod = $workbookPriceType.GetMethod("SaveAndPin")
$loadPriceMethod = $workbookPriceType.GetMethod("LoadRequired")
$loadRequirementsMethod = $requirementType.GetMethod("LoadPinnedNormRequirements")
$resolveSheetMethod = $roleServiceType.GetMethod("ResolveWorksheetRequired")
$createProfileMethod = $priceProfileType.GetMethod("Create")
$validateCoverageMethod = $coverageType.GetMethod("Validate")

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
$priceControl = $null
$shell = $null
$resourceSheet = $null
$landSheet = $null
try {
    $initialErrors = Get-ExcelErrorCount $workbook
    Invoke-Static $pinMethod @($workbook, $package2025) | Out-Null
    $resourceRole = [Enum]::Parse($worksheetRoleType, "ResourcePrices")
    $resourceSheet = Invoke-Static $resolveSheetMethod @($workbook, $resourceRole)
    $resourceSheet.Name = "DT404 Bang Gia"
    $landSheet = $workbook.Worksheets.Item("DG Can")

    $f7Formula = [string](Get-Cell $resourceSheet "F7" "Formula")
    $f23Formula = [string](Get-Cell $resourceSheet "F23" "Formula")
    $initialMaterial = [decimal](Get-Cell $resourceSheet "F81" "Value2")
    $dependentName = [string](Get-Cell $resourceSheet "B81" "Value2")
    $dependentRow = Find-RowByName $landSheet $dependentName
    $initialDependent = [decimal](Get-Cell $landSheet ("E" + $dependentRow) "Value2")

    $legacy = Invoke-Static $importLegacyMethod @($workbook, "Ha Noi", $null)
    if ($legacy.Entries.Count -ne 34) {
        throw "Import VL-NC-M phai co 34 entry, thuc te $($legacy.Entries.Count)."
    }
    if ([decimal]$legacy.FindRequired("MAT-CONCRETE-STAKE").AppliedUnitPriceVnd -ne 100000 -or
        [decimal]$legacy.FindRequired("bac-7-10").AppliedUnitPriceVnd -ne 495000 -or
        [decimal]$legacy.FindRequired("M011.001").AppliedUnitPriceVnd -ne 762996) {
        throw "Gia import legacy khong dung."
    }
    $requirements = @(Invoke-Static $loadRequirementsMethod @($workbook))
    $openListType = [Type]::GetType("System.Collections.Generic.List``1")
    $requirementListType = $openListType.MakeGenericType($priceRequirementType)
    $requirementList = [Activator]::CreateInstance($requirementListType)
    foreach ($requirement in $requirements) {
        [void]$requirementList.Add($requirement)
    }
    $coverageArguments = New-Object "object[]" 2
    $coverageArguments[0] = $legacy
    $coverageArguments[1] = $requirementList
    $coverage = $validateCoverageMethod.Invoke($null, $coverageArguments)
    if ($coverage.IsComplete -or $coverage.Issues.Count -lt 1) {
        throw "Workbook mau phai bao thieu gia thay vi tu gan 0."
    }

    $overrideArguments = New-Object "object[]" 7
    $overrideArguments[0] = "MAT-CONCRETE-STAKE"
    $overrideArguments[1] = [decimal]100000
    $overrideArguments[2] = [decimal]110000
    $overrideArguments[3] = "Bao gia moi"
    $overrideArguments[4] = "BG-DT404-01"
    $overrideArguments[5] = "IntegrationTest"
    $overrideArguments[6] = [DateTime]::UtcNow
    $priceOverride = [Activator]::CreateInstance($priceOverrideType, $overrideArguments)
    $overrideArray = [Array]::CreateInstance($priceOverrideType, 1)
    $overrideArray.SetValue($priceOverride, 0)
    $profileArguments = New-Object "object[]" 9
    $profileArguments[0] = $legacy.ProfileId
    $profileArguments[1] = "1.0.1"
    $profileArguments[2] = $legacy.DisplayName
    $profileArguments[3] = $legacy.Location
    $profileArguments[4] = $legacy.ValuationDate
    $profileArguments[5] = $legacy.LaborAudience
    $profileArguments[6] = $priceOverride.ModifiedAtUtc
    $profileArguments[7] = $legacy.Entries
    $profileArguments[8] = $overrideArray
    $overridden = $createProfileMethod.Invoke($null, $profileArguments)
    Invoke-Static $savePinMethod @($workbook, $overridden) | Out-Null
    $apply = Invoke-Static $applyLegacyMethod @($workbook, $overridden)
    if ($apply.Updated -ne 21 -or $apply.UnmatchedLegacyNames.Count -ne 0) {
        throw "Batch apply vat lieu khong du 21 dong."
    }
    if ([decimal](Get-Cell $resourceSheet "F81" "Value2") -ne 110000) {
        throw "VL-NC-M!F81 khong nhan override."
    }
    if ([decimal](Get-Cell $landSheet ("E" + $dependentRow) "Value2") -ne 110000) {
        throw "DG Can khong link gia moi tu sheet role."
    }
    if ([string](Get-Cell $resourceSheet "F7" "Formula") -ne $f7Formula -or
        [string](Get-Cell $resourceSheet "F23" "Formula") -ne $f23Formula) {
        throw "Apply vat lieu da ghi de cong thuc nhan cong/ca may."
    }

    $controlArguments = New-Object "object[]" 1
    $controlArguments[0] = $workbook.PSObject.BaseObject
    $priceControl = [Activator]::CreateInstance($controlType, $controlArguments)
    if ($priceControl.CurrentProfile.Entries.Count -ne 34 -or
        $priceControl.BuildCurrentProfile().Checksum -ne $overridden.Checksum) {
        throw "PriceProfileControl khong bao toan snapshot."
    }
    $profileIdInput = Get-PrivateField $controlType $priceControl "profileIdInput"
    $versionInput = Get-PrivateField $controlType $priceControl "versionInput"
    $displayNameInput = Get-PrivateField $controlType $priceControl "displayNameInput"
    $locationInput = Get-PrivateField $controlType $priceControl "locationInput"
    $audienceInput = Get-PrivateField $controlType $priceControl "audienceInput"
    $uiHost = New-Object System.Windows.Forms.Form
    $uiHost.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
    $uiHost.ShowInTaskbar = $false
    $uiHost.ClientSize = New-Object System.Drawing.Size(830, 614)
    $priceControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    [void]$uiHost.Controls.Add($priceControl)
    $uiHost.Show()
    [System.Windows.Forms.Application]::DoEvents()
    foreach ($metadataInput in @($profileIdInput, $versionInput, $displayNameInput, $locationInput)) {
        if (-not $metadataInput.Visible -or $metadataInput.Width -lt 20 -or
            $metadataInput.Height -lt 15 -or [string]::IsNullOrWhiteSpace($metadataInput.Text)) {
            throw "PriceProfile metadata input khong hien thi dung."
        }
    }
    if ($audienceInput.SelectedIndex -lt 0) {
        throw "PriceProfile audience chua duoc chon."
    }
    $bitmap = New-Object System.Drawing.Bitmap(830, 614)
    try {
        $priceControl.DrawToBitmap($bitmap, (New-Object System.Drawing.Rectangle(0, 0, 830, 614)))
        New-Item -ItemType Directory -Path (Split-Path -Parent $screenshotPath) -Force | Out-Null
        $bitmap.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bitmap.Dispose() }
    $uiHost.Controls.Remove($priceControl)
    $uiHost.Close()
    $uiHost.Dispose()
    $uiHost = $null

    $shellArguments = New-Object "object[]" 2
    $shellArguments[0] = $workbook.PSObject.BaseObject
    $shellArguments[1] = $null
    $shell = [Activator]::CreateInstance($shellType, $shellArguments)
    $rightPanel = Get-PrivateField $shellType $shell "pnlRight"
    $showPriceMethod = $shellType.GetMethod(
        "ShowPriceProfile",
        [Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)
    $showPriceMethod.Invoke($shell, (New-Object "object[]" 0)) | Out-Null
    [System.Windows.Forms.Application]::DoEvents()
    if ($rightPanel.Controls.Count -ne 1 -or
        $rightPanel.Controls[0].GetType().FullName -ne "ExcelAddIn1.Winform.PriceProfileControl") {
        throw "Shell Du toan khong mo PriceProfileControl."
    }

    $workbook.Save()
    $shell.Dispose(); $shell = $null
    $priceControl.Dispose(); $priceControl = $null
    Release-ComObject $landSheet; $landSheet = $null
    Release-ComObject $resourceSheet; $resourceSheet = $null
    $workbook.Close($false)
    Release-ComObject $workbook; $workbook = $null

    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $persistedOverride = Invoke-Static $loadPriceMethod @($workbook)
    if (-not $persistedOverride.FindRequired("MAT-CONCRETE-STAKE").IsOverridden) {
        throw "Override khong ton tai sau save/reopen."
    }
    $resourceSheet = Invoke-Static $resolveSheetMethod @($workbook, $resourceRole)
    if ($resourceSheet.Name -ne "DT404 Bang Gia" -or
        [decimal](Get-Cell $resourceSheet "F81" "Value2") -ne 110000) {
        throw "Role/price khong ton tai sau save/reopen."
    }

    Invoke-Static $savePinMethod @($workbook, $legacy) | Out-Null
    $restoreApply = Invoke-Static $applyLegacyMethod @($workbook, $legacy)
    if ($restoreApply.Updated -ne 21 -or
        [decimal](Get-Cell $resourceSheet "F81" "Value2") -ne $initialMaterial) {
        throw "Khong phuc hoi duoc gia baseline."
    }
    $workbook.Save()
    Release-ComObject $resourceSheet; $resourceSheet = $null
    $workbook.Close($false)
    Release-ComObject $workbook; $workbook = $null

    $verifyBook = $excel.Workbooks.Open($workingCopy, 0, $true)
    $persisted = Invoke-Static $loadPriceMethod @($verifyBook)
    $project = Invoke-Static $loadProjectMethod @($verifyBook)
    if ($persisted.Checksum -ne $legacy.Checksum -or
        $project.PriceProfileId -ne $legacy.ProfileId) {
        throw "Snapshot/profile pin sai sau phuc hoi."
    }
    $verifyResource = Invoke-Static $resolveSheetMethod @($verifyBook, $resourceRole)
    try {
        if ([decimal](Get-Cell $verifyResource "F81" "Value2") -ne $initialMaterial) {
            throw "Gia baseline sai sau reopen cuoi."
        }
    }
    finally { Release-ComObject $verifyResource }
    $verifySummary = $verifyBook.Worksheets.Item("THKP-TC")
    try { $summaryE27 = [double](Get-Cell $verifySummary "E27" "Value2") }
    finally { Release-ComObject $verifySummary }
    $finalErrors = Get-ExcelErrorCount $verifyBook
    if ($summaryE27 -ne 1198731000 -or $finalErrors -ne $initialErrors) {
        throw "Workbook regression sau phuc hoi khong dat."
    }

    [pscustomobject]@{
        Workbook = $workingCopy
        ImportedEntries = $legacy.Entries.Count
        MissingCoverageReported = $coverage.Issues.Count
        RenamedRoleSheet = "PASS"
        MaterialBatchUpdated = $apply.Updated
        LinkedUnitRate = "PASS"
        FormulaPreservation = "PASS"
        OverrideSaveReopen = "PASS"
        BaselineRestore = "PASS"
        UI = "PASS"
        Screenshot = $screenshotPath
        ExcelErrors = $finalErrors
        SummaryE27 = $summaryE27
    } | Format-List
}
finally {
    if ($null -ne $shell) { $shell.Dispose() }
    if ($null -ne $priceControl) { $priceControl.Dispose() }
    Release-ComObject $landSheet
    Release-ComObject $resourceSheet
    if ($null -ne $verifyBook) {
        try { $verifyBook.Close($false) } catch {}
        Release-ComObject $verifyBook
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    Stop-IsolatedExcelTestProcess $session
}
