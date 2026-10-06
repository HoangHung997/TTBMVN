param(
    [switch]$Overwrite
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot "excel-test-process.ps1")
$sourcePath = Join-Path $repoRoot "Dutoanmau\Checkpoints\DT-105-start.xlsm"
$variantDirectory = Join-Path $repoRoot "Dutoanmau\Variants"
$noTagsPath = Join-Path $variantDirectory "DT-105-no-tags.xlsm"
$missingPath = Join-Path $variantDirectory "DT-105-missing-sheet.xlsm"
$renamedPath = Join-Path $variantDirectory "DT-105-all-renamed.xlsm"

New-Item -ItemType Directory -Path $variantDirectory -Force | Out-Null
foreach ($path in @($noTagsPath, $missingPath, $renamedPath)) {
    if (Test-Path $path) {
        if (-not $Overwrite) {
            throw "Variant da ton tai: $path. Dung -Overwrite de tao lai."
        }
        Remove-Item -LiteralPath $path -Force
    }
    Copy-Item -LiteralPath $sourcePath -Destination $path
}

$corePath = Join-Path $repoRoot "ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll"
$addinPath = Join-Path $repoRoot "ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll"
$core = [Reflection.Assembly]::LoadFrom($corePath)
$addin = [Reflection.Assembly]::LoadFrom($addinPath)
$roleType = $core.GetType("ExcelAddIn1.Core.WorksheetRole", $true)
$mappingEntryType = $core.GetType("ExcelAddIn1.Core.WorksheetRoleMappingEntry", $true)
$profileType = $core.GetType("ExcelAddIn1.Core.ProjectProfile", $true)
$profileSchemaVersion = [int]$profileType.GetField("CurrentSchemaVersion").GetRawConstantValue()
$mappingListType = [System.Collections.Generic.List``1].MakeGenericType($mappingEntryType)
$roleCatalogType = $core.GetType("ExcelAddIn1.Core.WorksheetRoleCatalog", $true)
$mappingSuggesterType = $core.GetType("ExcelAddIn1.Core.WorksheetRoleMappingSuggester", $true)
$mappingValidatorType = $core.GetType("ExcelAddIn1.Core.WorksheetRoleMappingValidator", $true)
$roleServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorksheetRoleService", $true)
$profileServiceType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookProjectProfileService", $true)
$coordinatorType = $addin.GetType("ExcelAddIn1.Funtion.WorkbookSheetChangeCoordinator", $true)

$clearRoleMethod = $roleServiceType.GetMethod("ClearRole")
$validateRolesMethod = $roleServiceType.GetMethod("Validate")
$readAssignmentsMethod = $roleServiceType.GetMethod("ReadAssignments")
$applyMappingMethod = $roleServiceType.GetMethod("ApplyMapping")
$resolveRoleMethod = $roleServiceType.GetMethod("ResolveRequired")
$loadProfileMethod = $profileServiceType.GetMethod("LoadRequired")
$captureMethod = $coordinatorType.GetMethod("CaptureSnapshot")
$suggestMethod = $mappingSuggesterType.GetMethod("Suggest")
$validateMappingMethod = $mappingValidatorType.GetMethod("Validate")
$tryParseRoleMethod = $roleCatalogType.GetMethod("TryParse")

if (-not ("ReflectionInvocationBox" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Reflection;
using System.Runtime.CompilerServices;

public static class ReflectionInvocationBox
{
    private static object Unbox(object value)
    {
        StrongBox<object> box = value as StrongBox<object>;
        return box == null ? value : box.Value;
    }

    public static StrongBox<object> Invoke(MethodInfo method, object[] arguments)
    {
        if (method == null) throw new ArgumentNullException("method");
        return new StrongBox<object>(method.Invoke(null, arguments));
    }

    public static StrongBox<object> InvokeOne(MethodInfo method, object value)
    {
        return Invoke(method, new object[] { Unbox(value) });
    }

    public static StrongBox<object> InvokeTwo(MethodInfo method, object first, object second)
    {
        return Invoke(method, new object[] { Unbox(first), Unbox(second) });
    }

    public static StrongBox<object> Create(Type type)
    {
        return new StrongBox<object>(Activator.CreateInstance(type));
    }
}
"@
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Invoke-StaticMethod([Reflection.MethodInfo]$Method, [object[]]$Arguments) {
    try {
        $box = [ReflectionInvocationBox]::Invoke($Method, $Arguments)
    }
    catch {
        $types = @($Arguments | ForEach-Object {
            if ($null -eq $_) { "null" } else { $_.GetType().FullName }
        }) -join ", "
        throw "Invoke $($Method.DeclaringType.FullName).$($Method.Name) failed; args=[$types]. $($_.Exception.Message)"
    }
    $box
}

function Invoke-StaticOne([Reflection.MethodInfo]$Method, [object]$Value) {
    $argument = if ($null -eq $Value) { $null } else { $Value.PSObject.BaseObject }
    try {
        $box = [ReflectionInvocationBox]::InvokeOne($Method, $argument)
    }
    catch {
        $type = if ($null -eq $argument) { "null" } else { $argument.GetType().FullName }
        throw "Invoke $($Method.DeclaringType.FullName).$($Method.Name) failed; arg=[$type]. $($_.Exception.Message)"
    }
    $box
}

function Invoke-StaticTwo([Reflection.MethodInfo]$Method, [object]$First, [object]$Second) {
    $firstArgument = if ($null -eq $First) { $null } else { $First.PSObject.BaseObject }
    $secondArgument = if ($null -eq $Second) { $null } else { $Second.PSObject.BaseObject }
    try {
        $box = [ReflectionInvocationBox]::InvokeTwo($Method, $firstArgument, $secondArgument)
    }
    catch {
        $types = @($firstArgument, $secondArgument | ForEach-Object {
            if ($null -eq $_) { "null" } else { $_.GetType().FullName }
        }) -join ", "
        throw "Invoke $($Method.DeclaringType.FullName).$($Method.Name) failed; args=[$types]. $($_.Exception.Message)"
    }
    $box
}

function Clear-AllRoles([object]$Workbook) {
    $worksheets = $Workbook.Worksheets
    try {
        for ($index = 1; $index -le $worksheets.Count; $index++) {
            $worksheet = $worksheets.Item($index)
            try {
                [void](Invoke-StaticOne $clearRoleMethod $worksheet)
            }
            finally {
                Release-ComObject $worksheet
            }
        }
    }
    finally {
        Release-ComObject $worksheets
    }
}

function Validate-WorkbookState([object]$Workbook, [bool]$ExpectRoles) {
    $roles = Invoke-StaticOne $validateRolesMethod $Workbook
    if ($roles.Value.IsValid -ne $ExpectRoles) {
        throw "Role validation expected=$ExpectRoles actual=$($roles.Value.IsValid)."
    }
    $profile = Invoke-StaticOne $loadProfileMethod $Workbook
    if ($profile.Value.SchemaVersion -ne $profileSchemaVersion) {
        throw "ProjectProfile regression."
    }
    return $roles
}

function Get-E27([object]$Workbook) {
    $summaryRole = [Enum]::Parse($roleType, "CostSummary", $false)
    $summary = Invoke-StaticTwo $resolveRoleMethod $Workbook $summaryRole
    try {
        return [double]$summary.Value.Range("E27").Value2
    }
    finally {
        Release-ComObject $summary.Value
    }
}

$excelSession = Start-IsolatedExcelTestProcess
$excel = $excelSession.Application
$excel.EnableEvents = $false
$oldDisplayAlerts = $excel.DisplayAlerts
$results = @()
try {
    # Variant 1: remove every role marker, use first-run suggestions, apply and reopen.
    $workbook = $excel.Workbooks.Open($noTagsPath, 0, $false)
    try {
        Clear-AllRoles $workbook
        [void](Validate-WorkbookState $workbook $false)
        $sheets = Invoke-StaticOne $captureMethod $workbook
        $mapping = Invoke-StaticOne $suggestMethod $sheets
        $mappingValidation = Invoke-StaticTwo $validateMappingMethod $mapping $sheets
        if (-not $mappingValidation.Value.IsValid) {
            throw "First-run suggestion mapping khong valid."
        }
        $applied = Invoke-StaticTwo $applyMappingMethod $workbook $mapping
        if (-not $applied.Value.IsValid) {
            throw "Apply mapping no-tags khong valid."
        }
        $workbook.Save()
    }
    finally {
        $workbook.Close($false)
        Release-ComObject $workbook
    }

    $workbook = $excel.Workbooks.Open($noTagsPath, 0, $true)
    try {
        [void](Validate-WorkbookState $workbook $true)
        if ((Get-E27 $workbook) -ne 1198731000) {
            throw "No-tags E27 regression."
        }
        $results += [pscustomobject]@{ Variant = "no-tags"; Result = "PASS"; Sheets = $workbook.Worksheets.Count }
    }
    finally {
        $workbook.Close($false)
        Release-ComObject $workbook
    }

    # Variant 2: delete a required role sheet; validation and wizard draft must reject it.
    $workbook = $excel.Workbooks.Open($missingPath, 0, $false)
    try {
        $costRole = [Enum]::Parse($roleType, "CostRuleView", $false)
        $costSheet = Invoke-StaticTwo $resolveRoleMethod $workbook $costRole
        try {
            $excel.DisplayAlerts = $false
            $costSheet.Value.Delete()
        }
        finally {
            Release-ComObject $costSheet.Value
            $excel.DisplayAlerts = $oldDisplayAlerts
        }
        $roleResult = Validate-WorkbookState $workbook $false
        $sheets = Invoke-StaticOne $captureMethod $workbook
        $mapping = Invoke-StaticOne $suggestMethod $sheets
        $mappingValidation = Invoke-StaticTwo $validateMappingMethod $mapping $sheets
        if ($mappingValidation.Value.IsValid) {
            throw "Missing-sheet mapping phai bi tu choi."
        }
        $workbook.Save()
        $results += [pscustomobject]@{ Variant = "missing-sheet"; Result = "PASS-REJECTED"; Sheets = $workbook.Worksheets.Count }
    }
    finally {
        $workbook.Close($false)
        Release-ComObject $workbook
    }

    # Variant 3: rename every sheet, clear tags, then apply the user's selected mapping.
    $workbook = $excel.Workbooks.Open($renamedPath, 0, $false)
    try {
        $worksheets = $workbook.Worksheets
        try {
            for ($index = 1; $index -le $worksheets.Count; $index++) {
                $worksheet = $worksheets.Item($index)
                try {
                    $worksheet.Name = "S" + $index.ToString("D2") + "_DT105"
                }
                finally {
                    Release-ComObject $worksheet
                }
            }
        }
        finally {
            Release-ComObject $worksheets
        }

        [void](Validate-WorkbookState $workbook $true)
        $sheets = Invoke-StaticOne $captureMethod $workbook
        $assignments = Invoke-StaticOne $readAssignmentsMethod $workbook
        $manualMapping = [ReflectionInvocationBox]::Create($mappingListType)
        foreach ($assignment in $assignments.Value) {
            $roleArguments = New-Object "object[]" 2
            $roleArguments[0] = $assignment.RoleId
            $roleArguments[1] = [Enum]::ToObject($roleType, 0)
            if (-not [bool]$tryParseRoleMethod.Invoke($null, $roleArguments)) {
                throw "Unknown role before clear: $($assignment.RoleId)"
            }
            $role = $roleArguments[1]
            $descriptor = $sheets.Value | Where-Object { $_.Name -eq $assignment.SheetName } | Select-Object -First 1
            $entryArguments = New-Object "object[]" 3
            $entryArguments[0] = $role.PSObject.BaseObject
            $entryArguments[1] = $descriptor.Key
            $entryArguments[2] = $descriptor.Name
            $entry = [Activator]::CreateInstance($mappingEntryType, $entryArguments)
            [void]$manualMapping.Value.Add($entry)
        }
        Clear-AllRoles $workbook
        [void](Validate-WorkbookState $workbook $false)
        $mappingValidation = Invoke-StaticTwo $validateMappingMethod $manualMapping $sheets
        if (-not $mappingValidation.Value.IsValid) {
            throw "Manual all-renamed mapping khong valid."
        }
        [void](Invoke-StaticTwo $applyMappingMethod $workbook $manualMapping)
        [void](Validate-WorkbookState $workbook $true)
        $workbook.Save()
    }
    finally {
        $workbook.Close($false)
        Release-ComObject $workbook
    }

    $workbook = $excel.Workbooks.Open($renamedPath, 0, $true)
    try {
        [void](Validate-WorkbookState $workbook $true)
        $roleNames = [Enum]::GetNames($roleType)
        foreach ($roleName in $roleNames) {
            $role = [Enum]::Parse($roleType, $roleName, $false)
            $worksheet = Invoke-StaticTwo $resolveRoleMethod $workbook $role
            try {
                if ($worksheet.Value.Name -notlike "S*_DT105") {
                    throw "Role $roleName khong resolve den sheet da doi ten."
                }
            }
            finally {
                Release-ComObject $worksheet.Value
            }
        }
        if ((Get-E27 $workbook) -ne 1198731000) {
            throw "All-renamed E27 regression."
        }
        $results += [pscustomobject]@{ Variant = "all-renamed"; Result = "PASS"; Sheets = $workbook.Worksheets.Count }
    }
    finally {
        $workbook.Close($false)
        Release-ComObject $workbook
    }
}
finally {
    $excel.DisplayAlerts = $oldDisplayAlerts
    Stop-IsolatedExcelTestProcess $excelSession
}

$results | Format-Table -AutoSize
Get-FileHash $noTagsPath, $missingPath, $renamedPath -Algorithm SHA256 | Format-Table -AutoSize
