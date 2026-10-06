using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookPackageMigrationValueChange
    {
        internal WorkbookPackageMigrationValueChange(
            string key,
            string displayName,
            decimal sourceValue,
            decimal targetValue)
        {
            Key = key ?? string.Empty;
            DisplayName = displayName ?? string.Empty;
            SourceValue = sourceValue;
            TargetValue = targetValue;
        }

        public string Key { get; }
        public string DisplayName { get; }
        public decimal SourceValue { get; }
        public decimal TargetValue { get; }
        public decimal Difference => TargetValue - SourceValue;
    }

    public sealed class WorkbookPackageMigrationImpact
    {
        internal WorkbookPackageMigrationImpact(
            string priceProfileId,
            string priceProfileVersion,
            string priceProfileChecksum,
            int auditEntryCount,
            int estimateLineCount,
            IEnumerable<string> affectedScopes,
            IEnumerable<string> issues,
            IEnumerable<WorkbookPackageMigrationValueChange> valueChanges,
            string fingerprint)
        {
            PriceProfileId = priceProfileId ?? string.Empty;
            PriceProfileVersion = priceProfileVersion ?? string.Empty;
            PriceProfileChecksum = priceProfileChecksum ?? string.Empty;
            AuditEntryCount = auditEntryCount;
            EstimateLineCount = estimateLineCount;
            AffectedScopes = new ReadOnlyCollection<string>(
                (affectedScopes ?? Enumerable.Empty<string>()).Distinct(StringComparer.Ordinal).ToList());
            Issues = new ReadOnlyCollection<string>(
                (issues ?? Enumerable.Empty<string>()).Where(value => !string.IsNullOrWhiteSpace(value)).ToList());
            ValueChanges = new ReadOnlyCollection<WorkbookPackageMigrationValueChange>(
                (valueChanges ?? Enumerable.Empty<WorkbookPackageMigrationValueChange>()).ToList());
            Fingerprint = fingerprint ?? string.Empty;
        }

        public string PriceProfileId { get; }
        public string PriceProfileVersion { get; }
        public string PriceProfileChecksum { get; }
        public int AuditEntryCount { get; }
        public int EstimateLineCount { get; }
        public IReadOnlyList<string> AffectedScopes { get; }
        public IReadOnlyList<string> Issues { get; }
        public IReadOnlyList<WorkbookPackageMigrationValueChange> ValueChanges { get; }
        public string Fingerprint { get; }
        public bool CanApply => Issues.Count == 0;
    }

    public sealed class WorkbookPackageMigrationPreview
    {
        internal WorkbookPackageMigrationPreview(
            ProjectProfile sourceProfile,
            RegulationPackageMigrationPlan plan,
            WorkbookPackageMigrationImpact impact,
            WorkbookEstimatePreview targetEstimate,
            WorkbookCostSummaryPreview targetCostSummary,
            string workbookFingerprint)
        {
            SourceProfile = sourceProfile;
            Plan = plan;
            Impact = impact;
            TargetEstimate = targetEstimate;
            TargetCostSummary = targetCostSummary;
            WorkbookFingerprint = workbookFingerprint ?? string.Empty;
        }

        public ProjectProfile SourceProfile { get; }
        public RegulationPackageMigrationPlan Plan { get; }
        public WorkbookPackageMigrationImpact Impact { get; }
        internal WorkbookEstimatePreview TargetEstimate { get; }
        internal WorkbookCostSummaryPreview TargetCostSummary { get; }
        internal string WorkbookFingerprint { get; }
    }

    public sealed class WorkbookPackageMigrationResult
    {
        internal WorkbookPackageMigrationResult(
            WorkbookPackageMigrationStatus status,
            RegulationPackageMigrationPlan plan,
            WorkbookPackageMigrationImpact impact,
            string backupPath,
            WorkbookValidationReport validationReport)
        {
            Status = status;
            Plan = plan;
            Impact = impact;
            BackupPath = backupPath ?? string.Empty;
            ValidationReport = validationReport;
        }

        public WorkbookPackageMigrationStatus Status { get; }
        public RegulationPackageMigrationPlan Plan { get; }
        public WorkbookPackageMigrationImpact Impact { get; }
        public string BackupPath { get; }
        public WorkbookValidationReport ValidationReport { get; }
    }
}
