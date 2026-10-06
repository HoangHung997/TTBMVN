param(
    [string]$WorkbookPath = 'Y:\Mau DuToan\Du toan RPBM HoaLuNamDinh_Ver1.xlsx'
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'excel-test-process.ps1')
$runRoot = Join-Path $repoRoot ('tmp\V2-runtime-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$workingCopy = Join-Path $runRoot 'HoaLuNamDinh-V2-test.xlsx'
function Get-SharedHash([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','') }
    finally { $sha.Dispose(); $stream.Dispose() }
}
$sourceHash = Get-SharedHash $WorkbookPath
Copy-Item -LiteralPath $WorkbookPath -Destination $workingCopy
[void][Reflection.Assembly]::LoadFrom((Join-Path $repoRoot 'ExcelAddIn1.Core\bin\Release\ExcelAddIn1.Core.dll'))
$addin = [Reflection.Assembly]::LoadFrom((Join-Path $repoRoot 'ExcelAddIn1\bin\x64\Release\ExcelAddIn1.dll'))
function Invoke-V2([string]$TypeName, [string]$MethodName, [object[]]$Arguments) {
    $type = $addin.GetType('ExcelAddIn1.Funtion.' + $TypeName, $true)
    $method = $type.GetMethod($MethodName)
    $values = New-Object 'object[]' $Arguments.Length
    for ($i=0; $i -lt $Arguments.Length; $i++) {
        if ($null -ne $Arguments[$i]) { $values[$i] = $Arguments[$i].PSObject.BaseObject }
    }
    return $method.Invoke($null, $values)
}
function Assert-V2([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$session = Start-IsolatedExcelTestProcess
$excel = $session.Application
$excel.EnableEvents = $false
$excel.AutomationSecurity = 3
$workbook = $null
try {
    $workbook = $excel.Workbooks.Open($workingCopy, 0, $false)
    $workbook.Save()
    Invoke-V2 'WorkbookProjectProfileService' 'Clear' @($workbook) | Out-Null
    $available = Invoke-V2 'RegulationPackageBootstrapService' 'LoadAvailablePackages' @()
    $target = @($available | Where-Object PackageId -eq 'BQP-RPBM-2025')[0]
    Assert-V2 ($null -ne $target) 'Package BQP-RPBM-2025 unavailable'
    $preview = Invoke-V2 'WorkbookEstimateV2PackageService' 'Preview' @($workbook,$target,(Get-Date),(Get-Date))
    Assert-V2 $preview.Plan.CanApply 'Initial package preview blocked'
    $raw = New-Object 'object[]' 2
    $raw[0]=$workbook.PSObject.BaseObject
    $raw[1]=$null
    $hasProfile = $addin.GetType('ExcelAddIn1.Funtion.WorkbookProjectProfileService').GetMethod('TryLoad').Invoke($null,$raw)
    Assert-V2 (-not $hasProfile) 'Preview unexpectedly mutated profile'
    $canceled = Invoke-V2 'WorkbookEstimateV2PackageService' 'Apply' @($workbook,$preview,$false,$null)
    Assert-V2 ([string]::IsNullOrEmpty($canceled)) 'Cancel mutated workbook'
    $backup = Invoke-V2 'WorkbookEstimateV2PackageService' 'Apply' @($workbook,$preview,$true,$null)
    Assert-V2 (Test-Path -LiteralPath $backup) 'Package backup missing'
    $profile = Invoke-V2 'WorkbookProjectProfileService' 'LoadRequired' @($workbook)
    Assert-V2 ($profile.RegulationPackageChecksum -eq $target.PackageChecksum) 'Pinned identity mismatch'
    $beforePayload = $profile | ForEach-Object { [ExcelAddIn1.Core.ProjectProfileSerializer]::Serialize($_) }
    $again = Invoke-V2 'WorkbookEstimateV2PackageService' 'Preview' @($workbook,$target,(Get-Date),(Get-Date))
    $fail = [Action]{ throw 'V2 injected metadata failure' }
    $rolledBack=$false
    try { Invoke-V2 'WorkbookEstimateV2PackageService' 'Apply' @($workbook,$again,$true,$fail) | Out-Null }
    catch { $rolledBack=$true }
    Assert-V2 $rolledBack 'Failure injection did not fail'
    $restored = Invoke-V2 'WorkbookProjectProfileService' 'LoadRequired' @($workbook)
    Assert-V2 ([ExcelAddIn1.Core.ProjectProfileSerializer]::Serialize($restored) -eq $beforePayload) 'Rollback profile mismatch'
    $workbook.Save()
    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    $workbook=$excel.Workbooks.Open($workingCopy,0,$false)
    $reopened=Invoke-V2 'WorkbookProjectProfileService' 'LoadRequired' @($workbook)
    Assert-V2 ($reopened.RegulationPackageChecksum -eq $target.PackageChecksum) 'Save/reopen lost package'
    Assert-V2 ((Get-SharedHash $WorkbookPath) -eq $sourceHash) 'Source workbook changed'
    'PASS V2-801 preview/cancel/pin/backup/injected rollback/save-reopen'
    'WorkingCopy: ' + $workingCopy
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    Stop-IsolatedExcelTestProcess $session
}
