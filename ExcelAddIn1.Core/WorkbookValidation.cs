using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum WorkbookValidationSeverity
    {
        Info = 1,
        Warning = 2,
        Error = 3
    }

    public enum WorkbookValidationIssueCode
    {
        MissingPrice = 1,
        RoleMapping = 2,
        IncorrectTotal = 3,
        StaleFormula = 4,
        BrokenName = 5,
        StaleAudit = 6,
        InvalidProfile = 7,
        ScanFailure = 8
    }

    public sealed class WorkbookValidationIssue
    {
        public WorkbookValidationIssue(
            WorkbookValidationIssueCode code,
            WorkbookValidationSeverity severity,
            string worksheetRoleId,
            string worksheetCodeName,
            string address,
            string subject,
            string message,
            string remediation,
            bool isInherited = false)
        {
            Code = code;
            Severity = severity;
            WorksheetRoleId = Normalize(worksheetRoleId);
            WorksheetCodeName = Normalize(worksheetCodeName);
            Address = Normalize(address).ToUpperInvariant();
            Subject = Normalize(subject);
            Message = Normalize(message);
            Remediation = Normalize(remediation);
            IsInherited = isInherited;
        }

        public WorkbookValidationIssueCode Code { get; }
        public WorkbookValidationSeverity Severity { get; }
        public string WorksheetRoleId { get; }
        public string WorksheetCodeName { get; }
        public string Address { get; }
        public string Subject { get; }
        public string Message { get; }
        public string Remediation { get; }
        public bool IsInherited { get; }
        public bool IsBlocking => Severity == WorkbookValidationSeverity.Error && !IsInherited;

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim();
        }
    }

    public sealed class WorkbookValidationReport
    {
        public WorkbookValidationReport(IEnumerable<WorkbookValidationIssue> issues)
        {
            Issues = new ReadOnlyCollection<WorkbookValidationIssue>(
                (issues ?? Enumerable.Empty<WorkbookValidationIssue>())
                    .Where(issue => issue != null)
                    .OrderByDescending(issue => issue.Severity)
                    .ThenBy(issue => issue.Code)
                    .ThenBy(issue => issue.WorksheetRoleId, StringComparer.Ordinal)
                    .ThenBy(issue => issue.WorksheetCodeName, StringComparer.Ordinal)
                    .ThenBy(issue => issue.Address, StringComparer.Ordinal)
                    .ThenBy(issue => issue.Subject, StringComparer.Ordinal)
                    .ToList());
        }

        public IReadOnlyList<WorkbookValidationIssue> Issues { get; }
        public bool IsValid => Issues.All(issue => !issue.IsBlocking);
        public int ErrorCount => Issues.Count(issue => issue.IsBlocking);
        public int WarningCount => Issues.Count(issue =>
            issue.Severity == WorkbookValidationSeverity.Warning ||
            (issue.Severity == WorkbookValidationSeverity.Error && issue.IsInherited));

        public IReadOnlyList<WorkbookValidationIssue> Find(WorkbookValidationIssueCode code)
        {
            return new ReadOnlyCollection<WorkbookValidationIssue>(
                Issues.Where(issue => issue.Code == code).ToList());
        }
    }
}
