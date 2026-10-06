using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class SupportPackage
    {
        public static string Create()
        {
            return Create(null);
        }

        public static string Create(Excel.Workbook workbook)
        {
            Excel.Workbook activeWorkbook = null;
            bool ownsWorkbook = false;
            try
            {
                if (workbook == null)
                {
                    try
                    {
                        activeWorkbook = Globals.ThisAddIn?.Application?.ActiveWorkbook;
                        workbook = activeWorkbook;
                        ownsWorkbook = activeWorkbook != null;
                    }
                    catch
                    {
                    }
                }

                string stamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
                string folder = Path.Combine(
                    AppPaths.SupportDirectory,
                    "support_" + stamp + "_" + Guid.NewGuid().ToString("N").Substring(0, 8));
                Directory.CreateDirectory(folder);

                using (RuntimeLogScope operation = RuntimeLogger.BeginOperation(
                    "DT-604",
                    "support-package",
                    workbook,
                    string.Empty))
                {
                    WriteAppInfo(folder, workbook, operation.CorrelationId);
                    WriteActivationInfo(folder);
                    WriteWorkbookInfo(folder, workbook);
                    WriteProjectProfile(folder, workbook);
                    WritePackageManifest(folder, workbook);
                    WriteValidationReport(folder, workbook);
                    WriteCurrentSettings(folder);
                    WriteLog(folder, workbook);
                    WriteManifest(folder);
                    RuntimeLogger.LogMessage("Support package created", new Dictionary<string, string>
                    {
                        ["FileCount"] = Directory.GetFiles(folder).Length.ToString(CultureInfo.InvariantCulture)
                    });
                }
                return folder;
            }
            finally
            {
                if (ownsWorkbook)
                    Release(activeWorkbook);
            }
        }

        private static void WriteActivationInfo(string folder)
        {
            string path = Path.Combine(folder, "activation-info.txt");
            try
            {
                LicenseStatus license = LicenseManager.GetStatus();
                File.WriteAllLines(path, new[]
                {
                    "MachineToken: " + RuntimeLogger.GetMachineToken(),
                    "Activated: " + license.IsActivated,
                    "Trial: " + license.IsTrial,
                    "Expired: " + license.IsExpired,
                    "Expiry: " + (license.ExpiryDate.HasValue
                        ? license.ExpiryDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                        : "none"),
                    "TrialDaysRemaining: " + license.TrialDaysRemaining.ToString(CultureInfo.InvariantCulture)
                }, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                WriteError(path, ex);
            }
        }

        private static void WriteAppInfo(
            string folder,
            Excel.Workbook workbook,
            string correlationId)
        {
            LicenseStatus license = LicenseManager.GetStatus();
            var lines = new List<string>
            {
                "Product: " + AppInfo.ProductName,
                "Version: " + AppInfo.Version,
                "GeneratedUtc: " + DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                "CorrelationId: " + correlationId,
                "OS: " + Environment.OSVersion,
                ".NET: " + Environment.Version,
                "ProcessArchitecture: " + (Environment.Is64BitProcess ? "x64" : "x86"),
                "MachineToken: " + RuntimeLogger.GetMachineToken(),
                "LicenseActivated: " + license.IsActivated,
                "LicenseTrial: " + license.IsTrial,
                "LicenseExpired: " + license.IsExpired,
                "LicenseExpiry: " + (license.ExpiryDate.HasValue
                    ? license.ExpiryDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                    : "none"),
                "TrialDaysRemaining: " + license.TrialDaysRemaining.ToString(CultureInfo.InvariantCulture),
                "WorkbookId: " + RuntimeLogger.GetWorkbookToken(workbook)
            };

            try
            {
                Excel.Application application = workbook?.Application ?? Globals.ThisAddIn?.Application;
                if (application != null)
                {
                    lines.Add("ExcelVersion: " + application.Version);
                    lines.Add("ExcelCalculation: " + ((int)application.Calculation)
                        .ToString(CultureInfo.InvariantCulture));
                    lines.Add("ExcelDecimalSeparator: " + application.DecimalSeparator);
                    lines.Add("ExcelThousandsSeparator: " + application.ThousandsSeparator);
                }
            }
            catch (Exception ex)
            {
                lines.Add("ExcelInfoErrorType: " + ex.GetType().FullName);
                lines.Add("ExcelInfoErrorHResult: 0x" + ex.HResult.ToString("X8", CultureInfo.InvariantCulture));
            }

            File.WriteAllLines(Path.Combine(folder, "app-info.txt"), lines, Encoding.UTF8);
        }

        private static void WriteWorkbookInfo(string folder, Excel.Workbook workbook)
        {
            var lines = new List<string>
            {
                "WorkbookId: " + RuntimeLogger.GetWorkbookToken(workbook)
            };
            if (workbook == null)
            {
                lines.Add("WorkbookAvailable: False");
                File.WriteAllLines(Path.Combine(folder, "workbook-info.txt"), lines, Encoding.UTF8);
                return;
            }

            lines.Add("WorkbookAvailable: True");
            try
            {
                ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
                lines.Add("ProjectProfileSchema: " + profile.SchemaVersion.ToString(CultureInfo.InvariantCulture));
                lines.Add("PackageId: " + profile.RegulationPackageId);
                lines.Add("PackageVersion: " + profile.RegulationPackageVersion);
                lines.Add("PackageChecksum: " + profile.RegulationPackageChecksum);
                PriceProfile price = WorkbookPriceProfileService.LoadRequired(workbook);
                lines.Add("PriceProfileVersion: " + price.DataVersion);
                lines.Add("PriceProfileChecksum: " + price.Checksum);
                lines.Add("PriceEntryCount: " + price.Entries.Count.ToString(CultureInfo.InvariantCulture));
                ResultAuditTrail audit = WorkbookResultAuditService.Load(workbook);
                lines.Add("ResultAuditEntries: " + audit.Entries.Count.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                lines.Add("ProfileErrorType: " + ex.GetType().FullName);
                lines.Add("ProfileErrorHResult: 0x" + ex.HResult.ToString("X8", CultureInfo.InvariantCulture));
            }

            try
            {
                IReadOnlyList<WorksheetRoleAssignment> assignments =
                    WorksheetRoleService.ReadAssignments(workbook);
                WorksheetRoleValidationResult validation = WorksheetRoleValidator.Validate(assignments);
                lines.Add("WorksheetRoleValid: " + validation.IsValid);
                lines.Add("WorksheetRoleCount: " + assignments.Count.ToString(CultureInfo.InvariantCulture));
                foreach (IGrouping<string, WorksheetRoleAssignment> group in assignments
                    .GroupBy(item => item.RoleId, StringComparer.OrdinalIgnoreCase)
                    .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
                {
                    lines.Add("WorksheetRole." + group.Key + ": " +
                        group.Count().ToString(CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                lines.Add("RoleErrorType: " + ex.GetType().FullName);
                lines.Add("RoleErrorHResult: 0x" + ex.HResult.ToString("X8", CultureInfo.InvariantCulture));
            }
            File.WriteAllLines(Path.Combine(folder, "workbook-info.txt"), lines, Encoding.UTF8);
        }

        private static void WriteCurrentSettings(string folder)
        {
            string path = Path.Combine(folder, "daodat-settings.json");
            try
            {
                File.WriteAllText(
                    path,
                    RandomDaodat.GetSanitizedDiagnosticSettingsText(),
                    Encoding.UTF8);
            }
            catch (Exception ex)
            {
                WriteError(path, ex);
            }
        }

        private static void WriteProjectProfile(string folder, Excel.Workbook workbook)
        {
            string path = Path.Combine(folder, "project-profile.txt");
            try
            {
                if (workbook == null)
                {
                    File.WriteAllText(path, "WorkbookAvailable: False\n", Encoding.UTF8);
                    return;
                }
                ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
                string projectToken = Token("PRJ", profile.ProjectId);
                string priceProfileToken = Token("PRICE", profile.PriceProfileId);
                File.WriteAllLines(path, new[]
                {
                    "Schema: " + profile.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                    "ProjectToken: " + projectToken,
                    "PreparedDate: " + FormatDate(profile.PreparedDate),
                    "ApprovalDate: " + FormatDate(profile.ApprovalDate),
                    "PriceDate: " + FormatDate(profile.PriceDate),
                    "PackageId: " + profile.RegulationPackageId,
                    "PackageVersion: " + profile.RegulationPackageVersion,
                    "PackageChecksum: " + profile.RegulationPackageChecksum,
                    "PriceProfileToken: " + priceProfileToken,
                    "OverridePresent: " + !string.IsNullOrWhiteSpace(profile.OverrideSummary)
                }, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                WriteError(path, ex);
            }
        }

        private static void WritePackageManifest(string folder, Excel.Workbook workbook)
        {
            string path = Path.Combine(folder, "package-manifest.ttbmanifest");
            try
            {
                if (workbook == null)
                {
                    File.WriteAllText(path, "WorkbookAvailable: False\n", Encoding.UTF8);
                    return;
                }
                ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
                RegulationPackageBootstrapService.LoadAvailablePackages();
                var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
                RegulationPackage package = store.LoadRequired(
                    profile.RegulationPackageId,
                    profile.RegulationPackageVersion,
                    profile.RegulationPackageChecksum);
                File.WriteAllText(
                    path,
                    RegulationPackageSerializer.Serialize(package),
                    new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                WriteError(path, ex);
            }
        }

        private static void WriteValidationReport(string folder, Excel.Workbook workbook)
        {
            string path = Path.Combine(folder, "validation-report.txt");
            try
            {
                if (workbook == null)
                {
                    File.WriteAllText(path, "WorkbookAvailable: False\n", Encoding.UTF8);
                    return;
                }
                WorkbookValidationReport report = WorkbookValidationService.Scan(workbook);
                List<string> sensitive = CollectWorkbookSensitiveTokens(workbook);
                var lines = new List<string>
                {
                    "Valid: " + report.IsValid,
                    "BlockingErrors: " + report.ErrorCount.ToString(CultureInfo.InvariantCulture),
                    "Warnings: " + report.WarningCount.ToString(CultureInfo.InvariantCulture),
                    "IssueCount: " + report.Issues.Count.ToString(CultureInfo.InvariantCulture)
                };
                for (int index = 0; index < report.Issues.Count; index++)
                {
                    WorkbookValidationIssue issue = report.Issues[index];
                    string prefix = "Issue." + index.ToString(CultureInfo.InvariantCulture) + ".";
                    lines.Add(prefix + "Severity: " + issue.Severity);
                    lines.Add(prefix + "Code: " + issue.Code);
                    lines.Add(prefix + "Role: " + issue.WorksheetRoleId);
                    lines.Add(prefix + "Inherited: " + issue.IsInherited);
                    lines.Add(prefix + "Blocking: " + issue.IsBlocking);
                    lines.Add(prefix + "Subject: " + RuntimeLogger.SanitizeDiagnosticText(
                        issue.Subject, sensitive));
                    lines.Add(prefix + "Message: " + RuntimeLogger.SanitizeDiagnosticText(
                        issue.Message, sensitive));
                    lines.Add(prefix + "Remediation: " + RuntimeLogger.SanitizeDiagnosticText(
                        issue.Remediation, sensitive));
                }
                File.WriteAllLines(path, lines, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                WriteError(path, ex);
            }
        }

        private static void WriteLog(string folder, Excel.Workbook workbook)
        {
            try
            {
                var tokens = new List<string> { LicenseManager.GetMachineId() };
                if (workbook != null)
                {
                    try
                    {
                        tokens.Add(workbook.Name);
                        tokens.Add(workbook.FullName);
                    }
                    catch
                    {
                    }
                }
                RuntimeLogger.CreateSanitizedSnapshot(
                    Path.Combine(folder, "runtime.log"),
                    tokens);
            }
            catch (Exception ex)
            {
                WriteError(Path.Combine(folder, "runtime.log"), ex);
            }
        }

        private static List<string> CollectWorkbookSensitiveTokens(Excel.Workbook workbook)
        {
            var result = new List<string> { LicenseManager.GetMachineId() };
            if (workbook == null)
                return result;
            try
            {
                result.Add(workbook.Name);
                result.Add(workbook.FullName);
            }
            catch
            {
            }
            try
            {
                ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
                result.Add(profile.ProjectId);
                result.Add(profile.PriceProfileId);
            }
            catch
            {
            }
            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                foreach (Excel.Worksheet worksheet in worksheets)
                {
                    try
                    {
                        result.Add(worksheet.Name);
                        result.Add(worksheet.CodeName);
                    }
                    finally
                    {
                        Release(worksheet);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                Release(worksheets);
            }
            return result.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct().ToList();
        }

        private static void WriteManifest(string folder)
        {
            string manifestPath = Path.Combine(folder, "manifest.sha256");
            var lines = new List<string>();
            using (SHA256 sha = SHA256.Create())
            {
                foreach (string path in Directory.GetFiles(folder)
                    .Where(path => !string.Equals(path, manifestPath, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase))
                {
                    byte[] hash;
                    using (FileStream stream = File.OpenRead(path))
                        hash = sha.ComputeHash(stream);
                    lines.Add(ToHex(hash) + "  " + Path.GetFileName(path));
                }
            }
            File.WriteAllLines(manifestPath, lines, Encoding.ASCII);
        }

        private static string ToHex(byte[] value)
        {
            var builder = new StringBuilder(value.Length * 2);
            foreach (byte item in value)
                builder.Append(item.ToString("X2", CultureInfo.InvariantCulture));
            return builder.ToString();
        }

        private static void WriteError(string path, Exception exception)
        {
            File.WriteAllLines(
                path,
                new[]
                {
                    "Type: " + exception.GetType().FullName,
                    "HResult: 0x" + exception.HResult.ToString("X8", CultureInfo.InvariantCulture)
                },
                Encoding.UTF8);
        }

        private static string Token(string prefix, string value)
        {
            return prefix + "-" + ProjectProfileSerializer.ComputeChecksum(value ?? string.Empty)
                .Substring(0, 16);
        }

        private static string FormatDate(DateTime value)
        {
            return value == DateTime.MinValue
                ? "none"
                : value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static string FormatDate(DateTime? value)
        {
            return value.HasValue ? FormatDate(value.Value) : "none";
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
