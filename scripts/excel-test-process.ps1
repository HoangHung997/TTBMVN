$ErrorActionPreference = "Stop"

if (-not ("TTBMVN.Tests.ExcelNativeObject" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace TTBMVN.Tests
{
    public static class ExcelNativeObject
    {
        private const uint ObjIdNativeOm = 0xFFFFFFF0;

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd,
            uint objectId,
            ref Guid interfaceId,
            [MarshalAs(UnmanagedType.Interface)] out object value);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(
            IntPtr parent,
            IntPtr childAfter,
            string className,
            string windowTitle);

        private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr parameter);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(
            IntPtr parent,
            EnumWindowsProc callback,
            IntPtr parameter);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(
            IntPtr hwnd,
            StringBuilder className,
            int maximumCount);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(
            IntPtr hwnd,
            out uint processId);

        public static int GetProcessId(IntPtr mainWindowHandle)
        {
            uint processId;
            GetWindowThreadProcessId(mainWindowHandle, out processId);
            return checked((int)processId);
        }

        public static object GetApplication(IntPtr mainWindowHandle)
        {
            var desktopHandle = FindWindowEx(mainWindowHandle, IntPtr.Zero, "XLDESK", null);
            var nativeHandle = desktopHandle == IntPtr.Zero
                ? IntPtr.Zero
                : FindWindowEx(desktopHandle, IntPtr.Zero, "EXCEL7", null);
            if (nativeHandle == IntPtr.Zero)
                nativeHandle = FindWindowEx(mainWindowHandle, IntPtr.Zero, "EXCEL7", null);
            if (nativeHandle == IntPtr.Zero)
                nativeHandle = FindDescendant(mainWindowHandle, "EXCEL7");
            if (nativeHandle == IntPtr.Zero)
                throw new InvalidOperationException("Excel native document window was not found.");

            var dispatchId = new Guid("00020400-0000-0000-C000-000000000046");
            object nativeObject;
            var result = AccessibleObjectFromWindow(
                nativeHandle, ObjIdNativeOm, ref dispatchId, out nativeObject);
            if (result != 0 || nativeObject == null)
                Marshal.ThrowExceptionForHR(result);

            return nativeObject.GetType().InvokeMember(
                "Application",
                BindingFlags.GetProperty,
                null,
                nativeObject,
                null);
        }

        private static IntPtr FindDescendant(IntPtr parent, string expectedClass)
        {
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(parent, (hwnd, parameter) =>
            {
                var className = new StringBuilder(128);
                GetClassName(hwnd, className, className.Capacity);
                if (string.Equals(className.ToString(), expectedClass, StringComparison.Ordinal))
                {
                    found = hwnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }
    }
}
"@
}

function Get-ExcelExecutablePath {
    $running = Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.Path) } |
        Select-Object -First 1
    if ($null -ne $running) {
        return $running.Path
    }

    foreach ($key in @(
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")) {
        if (Test-Path -LiteralPath $key) {
            $value = (Get-Item -LiteralPath $key).GetValue("")
            if (-not [string]::IsNullOrWhiteSpace($value) -and (Test-Path -LiteralPath $value)) {
                return $value
            }
        }
    }
    throw "Khong tim thay EXCEL.EXE."
}

function Start-IsolatedExcelTestProcess([string]$InitialWorkbookPath) {
    $existingProcessIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Id })
    $application = $null
    try {
        $application = New-Object -ComObject Excel.Application
        $processId = [TTBMVN.Tests.ExcelNativeObject]::GetProcessId(
            [IntPtr]$application.Hwnd)
        if ($processId -le 0 -or $existingProcessIds -contains $processId) {
            throw "Excel COM factory khong tao tien trinh rieng."
        }
        $process = Get-Process -Id $processId
        $application.Visible = $false
        $application.DisplayAlerts = $false
        return [pscustomobject]@{
            Process = $process
            Application = $application
        }
    }
    catch {
        if ($null -ne $application) {
            try {
                $processId = [TTBMVN.Tests.ExcelNativeObject]::GetProcessId(
                    [IntPtr]$application.Hwnd)
                if ($processId -gt 0 -and $existingProcessIds -notcontains $processId) {
                    $application.Quit()
                }
            }
            catch {
            }
            if ([Runtime.InteropServices.Marshal]::IsComObject($application)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
            }
            $application = $null
        }
    }

    $excelPath = Get-ExcelExecutablePath
    $arguments = @("/x")
    if (-not [string]::IsNullOrWhiteSpace($InitialWorkbookPath)) {
        $arguments += ('"' + [IO.Path]::GetFullPath($InitialWorkbookPath) + '"')
    }
    # Hidden/minimized startup can omit the EXCEL7 native object used to bind
    # the exact /x process. Start normally, then hide immediately after binding.
    $process = Start-Process -FilePath $excelPath -ArgumentList $arguments `
        -WindowStyle Normal -PassThru
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    $application = $null
    try {
        while ([DateTime]::UtcNow -lt $deadline) {
            Start-Sleep -Milliseconds 250
            $process.Refresh()
            if ($process.HasExited) {
                throw "Tien trinh Excel test thoat truoc khi khoi tao COM."
            }
            if ($process.MainWindowHandle -eq [IntPtr]::Zero) {
                continue
            }
            try {
                $application = [TTBMVN.Tests.ExcelNativeObject]::GetApplication(
                    $process.MainWindowHandle)
                break
            }
            catch {
                # Main Excel window can exist briefly before the EXCEL7 document child.
            }
        }
        if ($null -eq $application) {
            throw "Khong lay duoc Excel Application cua tien trinh test PID $($process.Id)."
        }
        $application.Visible = $false
        $application.DisplayAlerts = $false
        return [pscustomobject]@{
            Process = $process
            Application = $application
        }
    }
    catch {
        if (-not $process.HasExited) {
            Stop-Process -Id $process.Id -Force
        }
        throw
    }
}

function Stop-IsolatedExcelTestProcess([object]$Session) {
    if ($null -eq $Session) { return }
    $application = $Session.Application
    if ($null -ne $application) {
        try { $application.Quit() } catch { }
        if ([Runtime.InteropServices.Marshal]::IsComObject($application)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $process = $Session.Process
    if ($null -ne $process) {
        try { [void]$process.WaitForExit(5000) } catch { }
        $process.Refresh()
        if (-not $process.HasExited) {
            Stop-Process -Id $process.Id -Force
        }
        $process.Dispose()
    }
}

function Get-OpenExcelWorkbookSession([string]$WorkbookPath) {
    $expectedPath = [IO.Path]::GetFullPath($WorkbookPath)
    foreach ($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        $application = $null
        $workbooks = $null
        $matched = $false
        try {
            $process.Refresh()
            if ($process.MainWindowHandle -eq [IntPtr]::Zero) { continue }
            $application = [TTBMVN.Tests.ExcelNativeObject]::GetApplication(
                $process.MainWindowHandle)
            $workbooks = $application.Workbooks
            for ($index = 1; $index -le $workbooks.Count; $index++) {
                $workbook = $workbooks.Item($index)
                if ([string]::Equals(
                    $workbook.FullName,
                    $expectedPath,
                    [StringComparison]::OrdinalIgnoreCase)) {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbooks)
                    $workbooks = $null
                    $matched = $true
                    return [pscustomobject]@{
                        Process = $process
                        Application = $application
                        Workbook = $workbook
                    }
                }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
            }
        }
        catch {
        }
        finally {
            if (-not $matched -and $null -ne $workbooks -and
                [Runtime.InteropServices.Marshal]::IsComObject($workbooks)) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbooks)
            }
            if (-not $matched -and $null -ne $application -and
                [Runtime.InteropServices.Marshal]::IsComObject($application)) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application)
            }
            if (-not $matched -and $null -ne $process) { $process.Dispose() }
        }
    }
    return $null
}

function Release-OpenExcelWorkbookSession([object]$Session) {
    if ($null -eq $Session) { return }
    foreach ($value in @($Session.Workbook, $Session.Application)) {
        if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
        }
    }
    if ($null -ne $Session.Process) { $Session.Process.Dispose() }
}
