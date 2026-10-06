using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public sealed class RegulationPackageResolutionRequest
    {
        public RegulationPackageResolutionRequest(
            DateTime preparedDate,
            DateTime? approvalDate,
            DateTime evaluationDate)
        {
            PreparedDate = preparedDate;
            ApprovalDate = approvalDate;
            EvaluationDate = evaluationDate;
        }

        public DateTime PreparedDate { get; }
        public DateTime? ApprovalDate { get; }
        public DateTime EvaluationDate { get; }
    }

    public sealed class RegulationTransitionRule
    {
        public RegulationTransitionRule(
            string ruleId,
            DateTime effectiveFrom,
            string previousPackageId,
            string newPackageId,
            string sourceDocumentId,
            string citation)
        {
            RuleId = (ruleId ?? string.Empty).Trim();
            EffectiveFrom = effectiveFrom;
            PreviousPackageId = (previousPackageId ?? string.Empty).Trim();
            NewPackageId = (newPackageId ?? string.Empty).Trim();
            SourceDocumentId = (sourceDocumentId ?? string.Empty).Trim();
            Citation = (citation ?? string.Empty).Trim();
        }

        public string RuleId { get; }
        public DateTime EffectiveFrom { get; }
        public string PreviousPackageId { get; }
        public string NewPackageId { get; }
        public string SourceDocumentId { get; }
        public string Citation { get; }
    }

    public enum RegulationPackageResolutionCode
    {
        SelectedByEvaluationDate,
        SelectedByApprovalDate,
        ApprovedBeforeTransition,
        NoApplicablePackage,
        AmbiguousPackages,
        RequiredTransitionPackageUnavailable,
        InvalidRequest
    }

    public sealed class RegulationPackageResolutionResult
    {
        internal RegulationPackageResolutionResult(
            RegulationPackageResolutionCode code,
            RegulationPackage package,
            DateTime? referenceDate,
            string ruleId,
            string reason,
            IEnumerable<string> sourceDocumentIds,
            IEnumerable<string> candidatePackageIds,
            IEnumerable<string> errors)
        {
            Code = code;
            Package = package;
            ReferenceDate = referenceDate;
            RuleId = ruleId ?? string.Empty;
            Reason = reason ?? string.Empty;
            SourceDocumentIds = CopyDistinct(sourceDocumentIds);
            CandidatePackageIds = CopyDistinct(candidatePackageIds);
            Errors = new ReadOnlyCollection<string>(
                (errors ?? Enumerable.Empty<string>()).ToList());
        }

        public bool IsSuccess => Package != null && Errors.Count == 0;
        public RegulationPackageResolutionCode Code { get; }
        public RegulationPackage Package { get; }
        public DateTime? ReferenceDate { get; }
        public string RuleId { get; }
        public string Reason { get; }
        public IReadOnlyList<string> SourceDocumentIds { get; }
        public IReadOnlyList<string> CandidatePackageIds { get; }
        public IReadOnlyList<string> Errors { get; }

        private static IReadOnlyList<string> CopyDistinct(IEnumerable<string> values)
        {
            return new ReadOnlyCollection<string>((values ?? Enumerable.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToList());
        }
    }

    public static class RegulationPackageResolver
    {
        public const string EffectiveDateRuleId = "EFFECTIVE-DATE-V1";

        public static RegulationPackageResolutionResult Resolve(
            RegulationPackageResolutionRequest request,
            IEnumerable<RegulationPackage> availablePackages,
            IEnumerable<RegulationTransitionRule> transitionRules)
        {
            var errors = ValidateRequest(request);
            List<RegulationPackage> packages = (availablePackages ??
                Enumerable.Empty<RegulationPackage>())
                .Where(package => package != null)
                .ToList();
            List<RegulationTransitionRule> rules = (transitionRules ??
                Enumerable.Empty<RegulationTransitionRule>())
                .Where(rule => rule != null)
                .ToList();
            ValidatePackages(packages, errors);
            ValidateRules(rules, errors);
            if (errors.Count > 0)
            {
                return new RegulationPackageResolutionResult(
                    RegulationPackageResolutionCode.InvalidRequest,
                    null,
                    null,
                    string.Empty,
                    "Khong the resolve package vi dau vao khong hop le.",
                    null,
                    null,
                    errors);
            }

            RegulationTransitionRule transition = rules
                .Where(rule =>
                    request.ApprovalDate.HasValue &&
                    request.ApprovalDate.Value.Date < rule.EffectiveFrom.Date &&
                    request.EvaluationDate.Date >= rule.EffectiveFrom.Date)
                .OrderByDescending(rule => rule.EffectiveFrom)
                .FirstOrDefault();
            if (transition != null)
            {
                RegulationPackage previous = packages
                    .Where(package =>
                        string.Equals(
                            package.PackageId,
                            transition.PreviousPackageId,
                            StringComparison.OrdinalIgnoreCase) &&
                        IsUsableHistoricalPackage(package) &&
                        IsEffectiveOn(package, request.ApprovalDate.Value.Date))
                    .OrderByDescending(package => package.EffectiveFrom)
                    .ThenByDescending(package => package.DataVersion, DataVersionComparer.Instance)
                    .FirstOrDefault();
                if (previous == null)
                {
                    return new RegulationPackageResolutionResult(
                        RegulationPackageResolutionCode.RequiredTransitionPackageUnavailable,
                        null,
                        request.ApprovalDate.Value.Date,
                        transition.RuleId,
                        "Ho so duoc phe duyet truoc moc chuyen tiep nhung package cu khong co trong kho.",
                        new[] { transition.SourceDocumentId },
                        new[] { transition.PreviousPackageId },
                        null);
                }

                return Success(
                    RegulationPackageResolutionCode.ApprovedBeforeTransition,
                    previous,
                    request.ApprovalDate.Value.Date,
                    transition.RuleId,
                    "Giu package cu vi phuong an va du toan da duoc phe duyet truoc ngay quy dinh moi co hieu luc.",
                    new[] { transition.SourceDocumentId });
            }

            DateTime referenceDate;
            RegulationPackageResolutionCode successCode;
            string reason;
            if (request.ApprovalDate.HasValue)
            {
                referenceDate = request.ApprovalDate.Value.Date;
                successCode = RegulationPackageResolutionCode.SelectedByApprovalDate;
                reason = "Chon package co hieu luc tai ngay phe duyet.";
            }
            else
            {
                referenceDate = request.EvaluationDate.Date;
                successCode = RegulationPackageResolutionCode.SelectedByEvaluationDate;
                reason = "Ho so chua phe duyet, chon package co hieu luc tai ngay tinh/kiem tra.";
            }

            List<RegulationPackage> candidates = CollapsePackagePatches(packages
                .Where(package =>
                    IsSelectablePackage(package) &&
                    IsEffectiveOn(package, referenceDate)))
                .OrderBy(package => package.PackageId, StringComparer.Ordinal)
                .ThenBy(package => package.DataVersion, StringComparer.Ordinal)
                .ToList();
            if (candidates.Count == 0)
            {
                return new RegulationPackageResolutionResult(
                    RegulationPackageResolutionCode.NoApplicablePackage,
                    null,
                    referenceDate,
                    EffectiveDateRuleId,
                    "Khong co package hop le tai ngay tham chieu.",
                    null,
                    null,
                    null);
            }
            if (candidates.Count > 1)
            {
                return new RegulationPackageResolutionResult(
                    RegulationPackageResolutionCode.AmbiguousPackages,
                    null,
                    referenceDate,
                    EffectiveDateRuleId,
                    "Nhieu package cung co hieu luc tai ngay tham chieu; can sua khoang hieu luc.",
                    null,
                    candidates.Select(FormatPackageIdentity),
                    null);
            }

            return Success(
                successCode,
                candidates[0],
                referenceDate,
                EffectiveDateRuleId,
                reason,
                null);
        }

        private static IEnumerable<RegulationPackage> CollapsePackagePatches(
            IEnumerable<RegulationPackage> packages)
        {
            return packages
                .GroupBy(package => package.PackageId, StringComparer.OrdinalIgnoreCase)
                .Select(group => group
                    .OrderByDescending(package => package.EffectiveFrom)
                    .ThenByDescending(package => package.DataVersion, DataVersionComparer.Instance)
                    .ThenByDescending(package => package.PackageChecksum, StringComparer.Ordinal)
                    .First());
        }

        private sealed class DataVersionComparer : IComparer<string>
        {
            internal static readonly DataVersionComparer Instance = new DataVersionComparer();

            public int Compare(string left, string right)
            {
                Version leftVersion;
                Version rightVersion;
                if (Version.TryParse(left, out leftVersion) && Version.TryParse(right, out rightVersion))
                    return leftVersion.CompareTo(rightVersion);
                return StringComparer.Ordinal.Compare(left, right);
            }
        }

        private static RegulationPackageResolutionResult Success(
            RegulationPackageResolutionCode code,
            RegulationPackage package,
            DateTime referenceDate,
            string ruleId,
            string reason,
            IEnumerable<string> additionalSourceIds)
        {
            IEnumerable<string> packageSources = package.Sources.Select(source => source.DocumentId);
            return new RegulationPackageResolutionResult(
                code,
                package,
                referenceDate,
                ruleId,
                reason,
                packageSources.Concat(additionalSourceIds ?? Enumerable.Empty<string>()),
                new[] { FormatPackageIdentity(package) },
                null);
        }

        private static List<string> ValidateRequest(RegulationPackageResolutionRequest request)
        {
            var errors = new List<string>();
            if (request == null)
            {
                errors.Add("Resolution request khong duoc null.");
                return errors;
            }
            ValidateDateOnly(request.PreparedDate, "PreparedDate", errors);
            ValidateDateOnly(request.EvaluationDate, "EvaluationDate", errors);
            if (request.ApprovalDate.HasValue)
            {
                ValidateDateOnly(request.ApprovalDate.Value, "ApprovalDate", errors);
                if (request.ApprovalDate.Value.Date < request.PreparedDate.Date)
                    errors.Add("ApprovalDate khong duoc truoc PreparedDate.");
                if (request.ApprovalDate.Value.Date > request.EvaluationDate.Date)
                    errors.Add("ApprovalDate khong duoc sau EvaluationDate.");
            }
            if (request.EvaluationDate.Date < request.PreparedDate.Date)
                errors.Add("EvaluationDate khong duoc truoc PreparedDate.");
            return errors;
        }

        private static void ValidatePackages(
            IEnumerable<RegulationPackage> packages,
            List<string> errors)
        {
            var identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (RegulationPackage package in packages)
            {
                RegulationPackageValidationResult validation = RegulationPackageValidator.Validate(package);
                if (!validation.IsValid)
                {
                    errors.Add("Package " + (package.PackageId ?? "<null>") + " khong hop le.");
                    continue;
                }
                string identity = package.PackageId + "|" + package.DataVersion + "|" + package.PackageChecksum;
                if (!identities.Add(identity))
                    errors.Add("Danh sach availablePackages bi trung package identity.");
            }
        }

        private static void ValidateRules(
            IEnumerable<RegulationTransitionRule> rules,
            List<string> errors)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (RegulationTransitionRule rule in rules)
            {
                if (rule.RuleId.Length == 0 || !ids.Add(rule.RuleId))
                    errors.Add("Transition RuleId trong hoac bi trung.");
                ValidateDateOnly(rule.EffectiveFrom, "Transition.EffectiveFrom", errors);
                if (rule.PreviousPackageId.Length == 0 || rule.NewPackageId.Length == 0)
                    errors.Add("Transition rule phai co previous/new package ID.");
                if (string.Equals(rule.PreviousPackageId, rule.NewPackageId, StringComparison.OrdinalIgnoreCase))
                    errors.Add("Transition previous/new package khong duoc trung nhau.");
                if (rule.SourceDocumentId.Length == 0 || rule.Citation.Length == 0)
                    errors.Add("Transition rule phai co can cu va citation.");
            }
        }

        private static bool IsSelectablePackage(RegulationPackage package)
        {
            return package.Status == RegulationPackageStatus.Published ||
                   package.Status == RegulationPackageStatus.Superseded;
        }

        private static bool IsUsableHistoricalPackage(RegulationPackage package)
        {
            return IsSelectablePackage(package);
        }

        private static string FormatPackageIdentity(RegulationPackage package)
        {
            return package.PackageId + "@" + package.DataVersion + "#" + package.PackageChecksum;
        }

        private static bool IsEffectiveOn(RegulationPackage package, DateTime date)
        {
            return package.EffectiveFrom.Date <= date.Date &&
                   (!package.EffectiveTo.HasValue || package.EffectiveTo.Value.Date >= date.Date);
        }

        private static void ValidateDateOnly(DateTime value, string fieldName, List<string> errors)
        {
            if (value.TimeOfDay != TimeSpan.Zero)
                errors.Add(fieldName + " chi duoc chua ngay, khong chua gio.");
        }
    }
}
