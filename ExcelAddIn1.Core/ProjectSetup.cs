using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public sealed class ProjectSetupDraft
    {
        public ProjectSetupDraft(
            string projectId,
            DateTime preparedDate,
            DateTime? approvalDate,
            DateTime evaluationDate,
            DateTime priceDate,
            string priceProfileId,
            string overrideSummary,
            IEnumerable<WorksheetRoleMappingEntry> sheetMapping)
        {
            ProjectId = (projectId ?? string.Empty).Trim();
            PreparedDate = preparedDate;
            ApprovalDate = approvalDate;
            EvaluationDate = evaluationDate;
            PriceDate = priceDate;
            PriceProfileId = (priceProfileId ?? string.Empty).Trim();
            OverrideSummary = overrideSummary ?? string.Empty;
            SheetMapping = new ReadOnlyCollection<WorksheetRoleMappingEntry>(
                (sheetMapping ?? Enumerable.Empty<WorksheetRoleMappingEntry>())
                    .Where(entry => entry != null)
                    .ToList());
        }

        public string ProjectId { get; }
        public DateTime PreparedDate { get; }
        public DateTime? ApprovalDate { get; }
        public DateTime EvaluationDate { get; }
        public DateTime PriceDate { get; }
        public string PriceProfileId { get; }
        public string OverrideSummary { get; }
        public IReadOnlyList<WorksheetRoleMappingEntry> SheetMapping { get; }
    }

    public sealed class ProjectSetupPlan
    {
        internal ProjectSetupPlan(
            ProjectProfile profile,
            RegulationPackageResolutionResult packageResolution,
            IEnumerable<WorksheetRoleMappingEntry> sheetMapping)
        {
            Profile = profile;
            PackageResolution = packageResolution;
            SheetMapping = new ReadOnlyCollection<WorksheetRoleMappingEntry>(
                (sheetMapping ?? Enumerable.Empty<WorksheetRoleMappingEntry>()).ToList());
        }

        public ProjectProfile Profile { get; }
        public RegulationPackageResolutionResult PackageResolution { get; }
        public IReadOnlyList<WorksheetRoleMappingEntry> SheetMapping { get; }
    }

    public sealed class ProjectSetupPlanningResult
    {
        internal ProjectSetupPlanningResult(
            ProjectSetupPlan plan,
            RegulationPackageResolutionResult packageResolution,
            IEnumerable<string> errors)
        {
            Plan = plan;
            PackageResolution = packageResolution;
            Errors = new ReadOnlyCollection<string>(
                (errors ?? Enumerable.Empty<string>()).ToList());
        }

        public bool IsValid => Plan != null && Errors.Count == 0;
        public ProjectSetupPlan Plan { get; }
        public RegulationPackageResolutionResult PackageResolution { get; }
        public IReadOnlyList<string> Errors { get; }
    }

    public static class ProjectSetupPlanner
    {
        public static ProjectSetupPlanningResult CreatePlan(
            ProjectSetupDraft draft,
            IEnumerable<WorkbookSheetDescriptor> availableSheets,
            IEnumerable<RegulationPackage> availablePackages,
            IEnumerable<RegulationTransitionRule> transitionRules)
        {
            var errors = new List<string>();
            if (draft == null)
            {
                errors.Add("Project setup draft khong duoc null.");
                return new ProjectSetupPlanningResult(null, null, errors);
            }

            RegulationPackageResolutionResult resolution = RegulationPackageResolver.Resolve(
                new RegulationPackageResolutionRequest(
                    draft.PreparedDate,
                    draft.ApprovalDate,
                    draft.EvaluationDate),
                availablePackages,
                transitionRules);
            if (!resolution.IsSuccess)
            {
                errors.AddRange(resolution.Errors);
                if (resolution.Errors.Count == 0)
                    errors.Add(resolution.Reason);
            }

            WorksheetRoleMappingValidationResult mappingValidation =
                WorksheetRoleMappingValidator.Validate(draft.SheetMapping, availableSheets);
            errors.AddRange(mappingValidation.Issues.Select(issue => issue.Message));

            ProjectProfile profile = null;
            if (resolution.IsSuccess)
            {
                profile = new ProjectProfile
                {
                    SchemaVersion = ProjectProfile.CurrentSchemaVersion,
                    ProjectId = draft.ProjectId,
                    PreparedDate = draft.PreparedDate.Date,
                    ApprovalDate = draft.ApprovalDate?.Date,
                    PriceDate = draft.PriceDate.Date,
                    RegulationPackageId = resolution.Package.PackageId,
                    RegulationPackageVersion = resolution.Package.DataVersion,
                    RegulationPackageChecksum = resolution.Package.PackageChecksum,
                    PriceProfileId = draft.PriceProfileId,
                    OverrideSummary = draft.OverrideSummary
                };
                ProjectProfileValidationResult profileValidation =
                    ProjectProfileValidator.Validate(profile);
                errors.AddRange(profileValidation.Errors);
            }

            if (errors.Count > 0)
                return new ProjectSetupPlanningResult(null, resolution, Distinct(errors));

            return new ProjectSetupPlanningResult(
                new ProjectSetupPlan(profile, resolution, draft.SheetMapping),
                resolution,
                new string[0]);
        }

        private static IEnumerable<string> Distinct(IEnumerable<string> errors)
        {
            return errors
                .Where(error => !string.IsNullOrWhiteSpace(error))
                .Select(error => error.Trim())
                .Distinct(StringComparer.Ordinal);
        }
    }

    public static class RegulationTransitionRuleCatalog
    {
        private static readonly ReadOnlyCollection<RegulationTransitionRule> Rules =
            Array.AsReadOnly(new[]
            {
                new RegulationTransitionRule(
                    "BQP-TT101-2025-ARTICLE-6",
                    new DateTime(2025, 10, 28),
                    "BQP-RPBM-2021",
                    "BQP-RPBM-2025",
                    "TT101-2025-BQP",
                    "Thong tu 101/2025/TT-BQP, Dieu 6")
            });

        public static IReadOnlyList<RegulationTransitionRule> All => Rules;
    }
}
