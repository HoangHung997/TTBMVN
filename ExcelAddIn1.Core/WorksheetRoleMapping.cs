using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public sealed class WorksheetRoleMappingEntry
    {
        public WorksheetRoleMappingEntry(WorksheetRole role, string sheetKey, string sheetName)
        {
            Role = role;
            SheetKey = (sheetKey ?? string.Empty).Trim();
            SheetName = (sheetName ?? string.Empty).Trim();
        }

        public WorksheetRole Role { get; }
        public string SheetKey { get; }
        public string SheetName { get; }
    }

    public enum WorksheetRoleMappingIssueCode
    {
        MissingRole,
        DuplicateRole,
        SheetNotSelected,
        SheetNotFound,
        SheetAssignedMultipleRoles
    }

    public sealed class WorksheetRoleMappingIssue
    {
        public WorksheetRoleMappingIssue(
            WorksheetRoleMappingIssueCode code,
            string message,
            IEnumerable<WorksheetRole> roles)
        {
            Code = code;
            Message = message ?? string.Empty;
            Roles = new ReadOnlyCollection<WorksheetRole>(
                (roles ?? Enumerable.Empty<WorksheetRole>()).Distinct().ToList());
        }

        public WorksheetRoleMappingIssueCode Code { get; }
        public string Message { get; }
        public IReadOnlyList<WorksheetRole> Roles { get; }
    }

    public sealed class WorksheetRoleMappingValidationResult
    {
        internal WorksheetRoleMappingValidationResult(IEnumerable<WorksheetRoleMappingIssue> issues)
        {
            Issues = new ReadOnlyCollection<WorksheetRoleMappingIssue>(
                (issues ?? Enumerable.Empty<WorksheetRoleMappingIssue>()).ToList());
        }

        public bool IsValid => Issues.Count == 0;
        public IReadOnlyList<WorksheetRoleMappingIssue> Issues { get; }
    }

    public static class WorksheetRoleMappingValidator
    {
        public static WorksheetRoleMappingValidationResult Validate(
            IEnumerable<WorksheetRoleMappingEntry> entries,
            IEnumerable<WorkbookSheetDescriptor> availableSheets)
        {
            var entryList = (entries ?? Enumerable.Empty<WorksheetRoleMappingEntry>())
                .Where(entry => entry != null)
                .ToList();
            var sheetList = WorkbookSheetList.Copy(availableSheets);
            var availableKeys = new HashSet<string>(
                sheetList.Select(sheet => sheet.Key),
                StringComparer.OrdinalIgnoreCase);
            var issues = new List<WorksheetRoleMappingIssue>();

            foreach (WorksheetRole role in WorksheetRoleCatalog.All)
            {
                int count = entryList.Count(entry => entry.Role == role);
                if (count == 0)
                {
                    issues.Add(new WorksheetRoleMappingIssue(
                        WorksheetRoleMappingIssueCode.MissingRole,
                        "Chua co dong anh xa cho vai tro " + WorksheetRoleCatalog.ToId(role) + ".",
                        new[] { role }));
                }
                else if (count > 1)
                {
                    issues.Add(new WorksheetRoleMappingIssue(
                        WorksheetRoleMappingIssueCode.DuplicateRole,
                        "Vai tro " + WorksheetRoleCatalog.ToId(role) + " co nhieu dong anh xa.",
                        new[] { role }));
                }
            }

            foreach (WorksheetRoleMappingEntry entry in entryList)
            {
                if (entry.SheetKey.Length == 0)
                {
                    issues.Add(new WorksheetRoleMappingIssue(
                        WorksheetRoleMappingIssueCode.SheetNotSelected,
                        "Chua chon sheet cho vai tro " + WorksheetRoleCatalog.ToId(entry.Role) + ".",
                        new[] { entry.Role }));
                }
                else if (!availableKeys.Contains(entry.SheetKey))
                {
                    issues.Add(new WorksheetRoleMappingIssue(
                        WorksheetRoleMappingIssueCode.SheetNotFound,
                        "Sheet da chon cho " + WorksheetRoleCatalog.ToId(entry.Role) + " khong con ton tai.",
                        new[] { entry.Role }));
                }
            }

            foreach (IGrouping<string, WorksheetRoleMappingEntry> group in entryList
                .Where(entry => entry.SheetKey.Length > 0)
                .GroupBy(entry => entry.SheetKey, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1))
            {
                issues.Add(new WorksheetRoleMappingIssue(
                    WorksheetRoleMappingIssueCode.SheetAssignedMultipleRoles,
                    "Mot sheet khong duoc gan cho nhieu vai tro.",
                    group.Select(entry => entry.Role)));
            }

            return new WorksheetRoleMappingValidationResult(issues);
        }
    }

    public static class WorksheetRoleMappingSuggester
    {
        private static readonly IReadOnlyDictionary<WorksheetRole, string> DefaultNames =
            new ReadOnlyDictionary<WorksheetRole, string>(
                new Dictionary<WorksheetRole, string>
                {
                    { WorksheetRole.ResourcePrices, "VL-NC-M" },
                    { WorksheetRole.UnitRateLand, "DG Can" },
                    { WorksheetRole.UnitRateWater, "DG Nuoc" },
                    { WorksheetRole.EstimateAppendix, "Gia DT TC" },
                    { WorksheetRole.CostSummary, "THKP-TC" },
                    { WorksheetRole.NormLookupView, "Tracuu" },
                    { WorksheetRole.CostRuleView, "ChiPhi" }
                });

        public static IReadOnlyList<WorksheetRoleMappingEntry> Suggest(
            IEnumerable<WorkbookSheetDescriptor> availableSheets)
        {
            IReadOnlyList<WorkbookSheetDescriptor> sheets = WorkbookSheetList.Copy(availableSheets);
            var result = new List<WorksheetRoleMappingEntry>();
            foreach (WorksheetRole role in WorksheetRoleCatalog.All)
            {
                string expectedName = DefaultNames[role];
                WorkbookSheetDescriptor match = sheets.FirstOrDefault(sheet =>
                    string.Equals(sheet.Name, expectedName, StringComparison.OrdinalIgnoreCase));
                result.Add(new WorksheetRoleMappingEntry(
                    role,
                    match?.Key ?? string.Empty,
                    match?.Name ?? string.Empty));
            }
            return result.AsReadOnly();
        }
    }

    public static class WorksheetRoleMappingBuilder
    {
        public static IReadOnlyList<WorksheetRoleMappingEntry> FromAssignments(
            IEnumerable<WorksheetRoleAssignment> assignments,
            IEnumerable<WorkbookSheetDescriptor> availableSheets)
        {
            IReadOnlyList<WorkbookSheetDescriptor> sheets =
                WorkbookSheetList.Copy(availableSheets);
            var selectedByRole = new Dictionary<WorksheetRole, WorkbookSheetDescriptor>();
            foreach (WorksheetRoleAssignment assignment in
                assignments ?? Enumerable.Empty<WorksheetRoleAssignment>())
            {
                if (assignment == null)
                    continue;
                WorksheetRole role;
                if (!WorksheetRoleCatalog.TryParse(assignment.RoleId, out role) ||
                    selectedByRole.ContainsKey(role))
                {
                    continue;
                }

                WorkbookSheetDescriptor sheet = sheets.FirstOrDefault(item =>
                    string.Equals(item.Key, assignment.SheetKey, StringComparison.OrdinalIgnoreCase)) ??
                    sheets.FirstOrDefault(item => string.Equals(
                        item.Name,
                        assignment.SheetName,
                        StringComparison.OrdinalIgnoreCase));
                if (sheet != null)
                    selectedByRole[role] = sheet;
            }

            IReadOnlyDictionary<WorksheetRole, WorksheetRoleMappingEntry> suggestions =
                new ReadOnlyDictionary<WorksheetRole, WorksheetRoleMappingEntry>(
                    WorksheetRoleMappingSuggester.Suggest(sheets)
                        .ToDictionary(entry => entry.Role));
            var result = new List<WorksheetRoleMappingEntry>();
            foreach (WorksheetRole role in WorksheetRoleCatalog.All)
            {
                WorkbookSheetDescriptor sheet;
                if (selectedByRole.TryGetValue(role, out sheet))
                {
                    result.Add(new WorksheetRoleMappingEntry(role, sheet.Key, sheet.Name));
                }
                else
                {
                    result.Add(suggestions[role]);
                }
            }
            return result.AsReadOnly();
        }
    }
}
