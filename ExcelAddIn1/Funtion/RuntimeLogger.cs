using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class RuntimeLogScope : IDisposable
    {
        private readonly RuntimeLogContext previous;
        private bool disposed;

        internal RuntimeLogScope(RuntimeLogContext context, RuntimeLogContext previousContext)
        {
            Context = context;
            previous = previousContext;
        }

        internal RuntimeLogContext Context { get; }
        public string CorrelationId => Context.CorrelationId;

        public void UpdatePhase(string phase)
        {
            Context.Phase = RuntimeLogger.NormalizeField(phase, "unspecified", 80);
        }

        public void Dispose()
        {
            if (disposed)
                return;
            disposed = true;
            RuntimeLogger.RestoreContext(Context, previous);
        }
    }

    internal sealed class RuntimeLogContext
    {
        internal string CorrelationId { get; set; }
        internal string Task { get; set; }
        internal string Phase { get; set; }
        internal string WorkbookId { get; set; }
        internal string WorksheetRole { get; set; }
        internal string PackageId { get; set; }
        internal string PackageVersion { get; set; }
        internal string PackageChecksum { get; set; }
        internal IReadOnlyList<string> SensitiveTokens { get; set; }
    }

    public static class RuntimeLogger
    {
        private const long MaxLogBytes = 5L * 1024L * 1024L;
        private const int ArchiveCount = 3;
        private static readonly object Sync = new object();
        private static readonly AsyncLocal<RuntimeLogContext> Current =
            new AsyncLocal<RuntimeLogContext>();
        private static readonly Regex WindowsPath = new Regex(
            @"(?i)(?:[a-z]:\\|\\\\)[^\r\n\t\""']+",
            RegexOptions.CultureInvariant);
        private static readonly Regex ExcelReference = new Regex(
            @"'[^'\r\n]{1,120}'!",
            RegexOptions.CultureInvariant);
        private static readonly Regex LicenseLikeToken = new Regex(
            @"\b[A-Z0-9]{4,8}(?:-[A-Z0-9]{4,8}){3,}\b",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public static string LogFilePath => AppPaths.LogFilePath;

        public static RuntimeLogScope BeginOperation(
            string task,
            string phase,
            Excel.Workbook workbook,
            string worksheetRoleId)
        {
            RuntimeLogContext previous = Current.Value;
            RuntimeLogContext context = CreateContext(
                task,
                phase,
                workbook,
                worksheetRoleId,
                previous?.CorrelationId);
            Current.Value = context;
            return new RuntimeLogScope(context, previous);
        }

        public static RuntimeLogScope BeginOperation(string task, string phase)
        {
            return BeginOperation(task, phase, null, string.Empty);
        }

        public static void Log(
            Exception exception,
            string action,
            IDictionary<string, string> context = null)
        {
            try
            {
                RuntimeLogContext operation = Current.Value ?? CreateContext(
                    "runtime",
                    "unspecified",
                    null,
                    string.Empty,
                    null);
                var builder = new StringBuilder();
                builder.AppendLine("==================================================");
                Field(builder, "TimestampUtc", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
                Field(builder, "CorrelationId", operation.CorrelationId);
                Field(builder, "Task", operation.Task);
                Field(builder, "Phase", operation.Phase);
                Field(builder, "WorkbookId", operation.WorkbookId);
                Field(builder, "WorksheetRole", operation.WorksheetRole);
                Field(builder, "PackageId", operation.PackageId);
                Field(builder, "PackageVersion", operation.PackageVersion);
                Field(builder, "PackageChecksum", operation.PackageChecksum);
                Field(builder, "Action", SanitizeText(action, operation.SensitiveTokens));

                if (context != null)
                {
                    foreach (KeyValuePair<string, string> item in context
                        .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
                    {
                        string key = NormalizeField(item.Key, "Unknown", 64);
                        string value = IsSensitiveKey(key)
                            ? "<redacted>"
                            : SanitizeText(item.Value, operation.SensitiveTokens);
                        Field(builder, "Context." + key, value);
                    }
                }

                AppendException(builder, exception, operation.SensitiveTokens);
                Write(builder.ToString());
            }
            catch
            {
                // Logging must never break the Excel workflow.
            }
        }

        public static void LogMessage(string action, IDictionary<string, string> context = null)
        {
            Log(null, action, context);
        }

        public static void LogOperation(
            Exception exception,
            string action,
            string task,
            string phase,
            Excel.Workbook workbook,
            string worksheetRoleId,
            IDictionary<string, string> context = null)
        {
            using (BeginOperation(task, phase, workbook, worksheetRoleId))
                Log(exception, action, context);
        }

        public static string GetWorkbookToken(Excel.Workbook workbook)
        {
            if (workbook == null)
                return "none";
            string identity;
            try
            {
                identity = workbook.FullName;
            }
            catch
            {
                try { identity = workbook.Name; }
                catch { identity = "unavailable"; }
            }
            return "WB-" + ProjectProfileSerializer.ComputeChecksum(
                (identity ?? string.Empty).Trim().ToUpperInvariant()).Substring(0, 16);
        }

        public static string GetMachineToken()
        {
            string machineId;
            try { machineId = LicenseManager.GetMachineId(); }
            catch { machineId = "unavailable"; }
            return "M-" + ProjectProfileSerializer.ComputeChecksum(machineId ?? string.Empty)
                .Substring(0, 16);
        }

        public static void CreateSanitizedSnapshot(
            string destinationPath,
            IEnumerable<string> sensitiveTokens = null)
        {
            if (string.IsNullOrWhiteSpace(destinationPath))
                throw new ArgumentException("Destination log path trong.", nameof(destinationPath));
            string content = string.Empty;
            lock (Sync)
            {
                if (File.Exists(LogFilePath))
                    content = File.ReadAllText(LogFilePath, Encoding.UTF8);
            }
            string sanitized = SanitizeText(content, sensitiveTokens);
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(destinationPath)));
            File.WriteAllText(destinationPath, sanitized, Encoding.UTF8);
        }

        public static string SanitizeDiagnosticText(
            string value,
            IEnumerable<string> sensitiveTokens = null)
        {
            return SanitizeText(value, sensitiveTokens);
        }

        internal static string NormalizeField(
            string value,
            string fallback,
            int maximumLength)
        {
            string normalized = (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ")
                .Trim();
            if (normalized.Length == 0)
                normalized = fallback;
            return normalized.Length <= maximumLength
                ? normalized
                : normalized.Substring(0, maximumLength);
        }

        internal static void RestoreContext(
            RuntimeLogContext expected,
            RuntimeLogContext previous)
        {
            if (ReferenceEquals(Current.Value, expected))
                Current.Value = previous;
        }

        private static RuntimeLogContext CreateContext(
            string task,
            string phase,
            Excel.Workbook workbook,
            string worksheetRoleId,
            string existingCorrelationId)
        {
            string packageId = "none";
            string packageVersion = "none";
            string packageChecksum = "none";
            var sensitive = new List<string>();
            if (workbook != null)
            {
                try
                {
                    sensitive.Add(workbook.Name);
                    sensitive.Add(workbook.FullName);
                }
                catch
                {
                }
                try
                {
                    ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
                    packageId = profile.RegulationPackageId;
                    packageVersion = profile.RegulationPackageVersion;
                    packageChecksum = profile.RegulationPackageChecksum;
                }
                catch
                {
                }
            }
            return new RuntimeLogContext
            {
                CorrelationId = string.IsNullOrWhiteSpace(existingCorrelationId)
                    ? Guid.NewGuid().ToString("N")
                    : existingCorrelationId,
                Task = NormalizeField(task, "runtime", 40),
                Phase = NormalizeField(phase, "unspecified", 80),
                WorkbookId = GetWorkbookToken(workbook),
                WorksheetRole = NormalizeField(worksheetRoleId, "none", 80),
                PackageId = NormalizeField(packageId, "none", 160),
                PackageVersion = NormalizeField(packageVersion, "none", 40),
                PackageChecksum = NormalizeChecksum(packageChecksum),
                SensitiveTokens = sensitive
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderByDescending(value => value.Length)
                    .ToArray()
            };
        }

        private static string NormalizeChecksum(string value)
        {
            string text = (value ?? string.Empty).Trim().ToUpperInvariant();
            return text.Length == 64 && text.All(IsHexDigit) ? text : "none";
        }

        private static bool IsHexDigit(char value)
        {
            return (value >= '0' && value <= '9') ||
                (value >= 'A' && value <= 'F');
        }

        private static void AppendException(
            StringBuilder builder,
            Exception exception,
            IEnumerable<string> sensitiveTokens)
        {
            int depth = 0;
            for (Exception current = exception; current != null && depth < 8; current = current.InnerException)
            {
                string prefix = depth == 0 ? "Exception" : "InnerException" + depth;
                Field(builder, prefix + ".Type", current.GetType().FullName);
                Field(builder, prefix + ".HResult", "0x" + current.HResult.ToString("X8", CultureInfo.InvariantCulture));
                Field(builder, prefix + ".Message", SanitizeText(current.Message, sensitiveTokens));
                if (!string.IsNullOrWhiteSpace(current.StackTrace))
                    Field(builder, prefix + ".StackTrace", SanitizeText(current.StackTrace, sensitiveTokens));
                depth++;
            }
        }

        private static bool IsSensitiveKey(string key)
        {
            string value = key ?? string.Empty;
            return value.IndexOf("path", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("name", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("key", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("value", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("address", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("data", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("license", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("machine", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("customer", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("project", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string SanitizeText(
            string value,
            IEnumerable<string> sensitiveTokens)
        {
            string text = value ?? string.Empty;
            if (sensitiveTokens != null)
            {
                foreach (string token in sensitiveTokens
                    .Where(token => !string.IsNullOrWhiteSpace(token))
                    .OrderByDescending(token => token.Length))
                {
                    text = text.Replace(token, "<redacted>");
                }
            }
            text = WindowsPath.Replace(text, "<path>");
            text = ExcelReference.Replace(text, "'<sheet>'!");
            text = LicenseLikeToken.Replace(text, "<token>");
            return text.Length <= 32768 ? text : text.Substring(0, 32768) + "<truncated>";
        }

        private static void Field(StringBuilder builder, string key, string value)
        {
            builder.Append(key).Append(": ")
                .Append((value ?? string.Empty).Replace("\r", "\\r").Replace("\n", "\\n"))
                .AppendLine();
        }

        private static void Write(string text)
        {
            lock (Sync)
            {
                string path = LogFilePath;
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                RotateIfRequired(path, Encoding.UTF8.GetByteCount(text));
                File.AppendAllText(path, text, Encoding.UTF8);
            }
        }

        private static void RotateIfRequired(string path, int incomingBytes)
        {
            if (!File.Exists(path) || new FileInfo(path).Length + incomingBytes <= MaxLogBytes)
                return;
            string oldest = path + "." + ArchiveCount.ToString(CultureInfo.InvariantCulture);
            if (File.Exists(oldest))
                File.Delete(oldest);
            for (int index = ArchiveCount - 1; index >= 1; index--)
            {
                string source = path + "." + index.ToString(CultureInfo.InvariantCulture);
                string target = path + "." + (index + 1).ToString(CultureInfo.InvariantCulture);
                if (File.Exists(source))
                    File.Move(source, target);
            }
            File.Move(path, path + ".1");
        }
    }
}
