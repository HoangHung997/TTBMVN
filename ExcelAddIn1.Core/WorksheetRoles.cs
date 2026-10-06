using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum WorksheetRole
    {
        ResourcePrices,
        UnitRateLand,
        UnitRateWater,
        EstimateAppendix,
        CostSummary,
        NormLookupView,
        CostRuleView
    }

    public static class WorksheetRoleCatalog
    {
        private static readonly ReadOnlyCollection<WorksheetRole> Roles =
            Array.AsReadOnly(new[]
            {
                WorksheetRole.ResourcePrices,
                WorksheetRole.UnitRateLand,
                WorksheetRole.UnitRateWater,
                WorksheetRole.EstimateAppendix,
                WorksheetRole.CostSummary,
                WorksheetRole.NormLookupView,
                WorksheetRole.CostRuleView
            });

        public static IReadOnlyList<WorksheetRole> All => Roles;

        public static string ToId(WorksheetRole role)
        {
            if (!Roles.Contains(role))
                throw new ArgumentOutOfRangeException(nameof(role), role, "Vai tro sheet khong hop le.");

            return role.ToString();
        }

        public static bool TryParse(string roleId, out WorksheetRole role)
        {
            role = default(WorksheetRole);
            string normalized = (roleId ?? string.Empty).Trim();
            if (normalized.Length == 0)
                return false;

            foreach (WorksheetRole candidate in Roles)
            {
                if (string.Equals(ToId(candidate), normalized, StringComparison.OrdinalIgnoreCase))
                {
                    role = candidate;
                    return true;
                }
            }

            return false;
        }
    }

    public sealed class WorksheetRoleAssignment
    {
        public WorksheetRoleAssignment(string sheetKey, string sheetName, string roleId)
        {
            SheetKey = (sheetKey ?? string.Empty).Trim();
            SheetName = (sheetName ?? string.Empty).Trim();
            RoleId = (roleId ?? string.Empty).Trim();
        }

        public string SheetKey { get; }
        public string SheetName { get; }
        public string RoleId { get; }
    }

    public enum WorksheetRoleIssueCode
    {
        MissingRole,
        DuplicateRole,
        UnknownRole,
        DuplicateSheet
    }

    public sealed class WorksheetRoleValidationIssue
    {
        public WorksheetRoleValidationIssue(
            WorksheetRoleIssueCode code,
            string roleId,
            IEnumerable<string> sheetNames,
            string message)
        {
            Code = code;
            RoleId = roleId ?? string.Empty;
            SheetNames = new ReadOnlyCollection<string>((sheetNames ?? Enumerable.Empty<string>()).ToList());
            Message = message ?? string.Empty;
        }

        public WorksheetRoleIssueCode Code { get; }
        public string RoleId { get; }
        public IReadOnlyList<string> SheetNames { get; }
        public string Message { get; }
    }

    public sealed class WorksheetRoleValidationResult
    {
        internal WorksheetRoleValidationResult(IEnumerable<WorksheetRoleValidationIssue> issues)
        {
            Issues = new ReadOnlyCollection<WorksheetRoleValidationIssue>(issues.ToList());
        }

        public bool IsValid => Issues.Count == 0;
        public IReadOnlyList<WorksheetRoleValidationIssue> Issues { get; }
    }

    public static class WorksheetRoleValidator
    {
        public static WorksheetRoleValidationResult Validate(
            IEnumerable<WorksheetRoleAssignment> assignments,
            IEnumerable<WorksheetRole> requiredRoles = null)
        {
            var assignmentList = (assignments ?? Enumerable.Empty<WorksheetRoleAssignment>())
                .Where(item => item != null)
                .ToList();
            var issues = new List<WorksheetRoleValidationIssue>();

            foreach (IGrouping<string, WorksheetRoleAssignment> group in assignmentList
                .Where(item => item.SheetKey.Length > 0)
                .GroupBy(item => item.SheetKey, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1))
            {
                issues.Add(new WorksheetRoleValidationIssue(
                    WorksheetRoleIssueCode.DuplicateSheet,
                    string.Empty,
                    group.Select(item => DisplaySheet(item)),
                    "Mot sheet co nhieu khai bao vai tro."));
            }

            var knownAssignments = new List<Tuple<WorksheetRoleAssignment, WorksheetRole>>();
            foreach (WorksheetRoleAssignment assignment in assignmentList)
            {
                WorksheetRole role;
                if (!WorksheetRoleCatalog.TryParse(assignment.RoleId, out role))
                {
                    issues.Add(new WorksheetRoleValidationIssue(
                        WorksheetRoleIssueCode.UnknownRole,
                        assignment.RoleId,
                        new[] { DisplaySheet(assignment) },
                        "Khong nhan biet vai tro sheet: " + assignment.RoleId));
                    continue;
                }

                knownAssignments.Add(Tuple.Create(assignment, role));
            }

            foreach (IGrouping<WorksheetRole, Tuple<WorksheetRoleAssignment, WorksheetRole>> group in
                knownAssignments.GroupBy(item => item.Item2).Where(group => group.Count() > 1))
            {
                string roleId = WorksheetRoleCatalog.ToId(group.Key);
                issues.Add(new WorksheetRoleValidationIssue(
                    WorksheetRoleIssueCode.DuplicateRole,
                    roleId,
                    group.Select(item => DisplaySheet(item.Item1)),
                    "Vai tro " + roleId + " duoc gan cho nhieu sheet."));
            }

            IEnumerable<WorksheetRole> rolesToRequire = requiredRoles ?? WorksheetRoleCatalog.All;
            foreach (WorksheetRole requiredRole in rolesToRequire.Distinct())
            {
                if (!knownAssignments.Any(item => item.Item2 == requiredRole))
                {
                    string roleId = WorksheetRoleCatalog.ToId(requiredRole);
                    issues.Add(new WorksheetRoleValidationIssue(
                        WorksheetRoleIssueCode.MissingRole,
                        roleId,
                        new string[0],
                        "Chua gan sheet cho vai tro " + roleId + "."));
                }
            }

            return new WorksheetRoleValidationResult(issues);
        }

        private static string DisplaySheet(WorksheetRoleAssignment assignment)
        {
            if (assignment.SheetName.Length > 0)
                return assignment.SheetName;
            if (assignment.SheetKey.Length > 0)
                return assignment.SheetKey;
            return "(khong co dinh danh)";
        }
    }

    public static class WorksheetRoleResolver
    {
        public static WorksheetRoleAssignment ResolveUnique(
            IEnumerable<WorksheetRoleAssignment> assignments,
            WorksheetRole role)
        {
            string expectedRoleId = WorksheetRoleCatalog.ToId(role);
            var matches = (assignments ?? Enumerable.Empty<WorksheetRoleAssignment>())
                .Where(item => item != null)
                .Where(item =>
                {
                    WorksheetRole parsed;
                    return WorksheetRoleCatalog.TryParse(item.RoleId, out parsed) && parsed == role;
                })
                .ToList();

            if (matches.Count == 0)
                throw new InvalidOperationException("Khong tim thay sheet co vai tro " + expectedRoleId + ".");
            if (matches.Count > 1)
                throw new InvalidOperationException("Vai tro " + expectedRoleId + " dang bi gan trung cho nhieu sheet.");

            return matches[0];
        }
    }
}

