using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum RegulationPackageChangeKind
    {
        PackageMetadataChanged,
        SourceAdded,
        SourceRemoved,
        SourceChanged,
        ModuleAdded,
        ModuleRemoved,
        ModuleChanged
    }

    public enum RegulationPackageChangeImpact
    {
        Informational,
        LegalBasis,
        CalculationData
    }

    public sealed class RegulationPackageChange
    {
        public RegulationPackageChange(
            RegulationPackageChangeKind kind,
            RegulationPackageChangeImpact impact,
            string key,
            string oldValue,
            string newValue,
            RegulationModuleKind? moduleKind)
        {
            Kind = kind;
            Impact = impact;
            Key = key ?? string.Empty;
            OldValue = oldValue ?? string.Empty;
            NewValue = newValue ?? string.Empty;
            ModuleKind = moduleKind;
        }

        public RegulationPackageChangeKind Kind { get; }
        public RegulationPackageChangeImpact Impact { get; }
        public string Key { get; }
        public string OldValue { get; }
        public string NewValue { get; }
        public RegulationModuleKind? ModuleKind { get; }
    }

    public sealed class RegulationPackageDiffResult
    {
        internal RegulationPackageDiffResult(IEnumerable<RegulationPackageChange> changes)
        {
            Changes = new ReadOnlyCollection<RegulationPackageChange>(
                (changes ?? Enumerable.Empty<RegulationPackageChange>()).ToList());
        }

        public bool HasChanges => Changes.Count > 0;
        public IReadOnlyList<RegulationPackageChange> Changes { get; }
        public int LegalBasisChangeCount => Changes.Count(change =>
            change.Impact == RegulationPackageChangeImpact.LegalBasis);
        public int CalculationDataChangeCount => Changes.Count(change =>
            change.Impact == RegulationPackageChangeImpact.CalculationData);
    }

    public static class RegulationPackageDiffer
    {
        public static RegulationPackageDiffResult Compare(
            RegulationPackage source,
            RegulationPackage target)
        {
            ValidatePackage(source, nameof(source));
            ValidatePackage(target, nameof(target));
            var changes = new List<RegulationPackageChange>();

            AddMetadataChange(changes, "PackageId", source.PackageId, target.PackageId);
            AddMetadataChange(changes, "DataVersion", source.DataVersion, target.DataVersion);
            AddMetadataChange(
                changes,
                "EffectiveFrom",
                FormatDate(source.EffectiveFrom),
                FormatDate(target.EffectiveFrom));
            AddMetadataChange(
                changes,
                "EffectiveTo",
                FormatDate(source.EffectiveTo),
                FormatDate(target.EffectiveTo));
            AddMetadataChange(changes, "Status", source.Status.ToString(), target.Status.ToString());
            AddMetadataChange(changes, "TransitionNote", source.TransitionNote, target.TransitionNote);

            CompareSources(source.Sources, target.Sources, changes);
            CompareModules(source.Modules, target.Modules, changes);
            return new RegulationPackageDiffResult(changes
                .OrderBy(change => change.Kind)
                .ThenBy(change => change.Key, StringComparer.Ordinal));
        }

        private static void CompareSources(
            IEnumerable<RegulationPackageSourceDocument> sourceItems,
            IEnumerable<RegulationPackageSourceDocument> targetItems,
            List<RegulationPackageChange> changes)
        {
            Dictionary<string, RegulationPackageSourceDocument> source = sourceItems
                .ToDictionary(item => item.DocumentId, StringComparer.OrdinalIgnoreCase);
            Dictionary<string, RegulationPackageSourceDocument> target = targetItems
                .ToDictionary(item => item.DocumentId, StringComparer.OrdinalIgnoreCase);
            foreach (string id in source.Keys.Union(target.Keys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.Ordinal))
            {
                RegulationPackageSourceDocument oldItem;
                RegulationPackageSourceDocument newItem;
                bool hasOld = source.TryGetValue(id, out oldItem);
                bool hasNew = target.TryGetValue(id, out newItem);
                if (!hasOld)
                {
                    changes.Add(new RegulationPackageChange(
                        RegulationPackageChangeKind.SourceAdded,
                        RegulationPackageChangeImpact.LegalBasis,
                        id,
                        string.Empty,
                        DescribeSource(newItem),
                        null));
                }
                else if (!hasNew)
                {
                    changes.Add(new RegulationPackageChange(
                        RegulationPackageChangeKind.SourceRemoved,
                        RegulationPackageChangeImpact.LegalBasis,
                        id,
                        DescribeSource(oldItem),
                        string.Empty,
                        null));
                }
                else
                {
                    string oldValue = DescribeSource(oldItem);
                    string newValue = DescribeSource(newItem);
                    if (!string.Equals(oldValue, newValue, StringComparison.Ordinal))
                    {
                        changes.Add(new RegulationPackageChange(
                            RegulationPackageChangeKind.SourceChanged,
                            RegulationPackageChangeImpact.LegalBasis,
                            id,
                            oldValue,
                            newValue,
                            null));
                    }
                }
            }
        }

        private static void CompareModules(
            IEnumerable<RegulationPackageModuleManifest> sourceItems,
            IEnumerable<RegulationPackageModuleManifest> targetItems,
            List<RegulationPackageChange> changes)
        {
            Dictionary<RegulationModuleKind, RegulationPackageModuleManifest> source =
                sourceItems.ToDictionary(item => item.Kind);
            Dictionary<RegulationModuleKind, RegulationPackageModuleManifest> target =
                targetItems.ToDictionary(item => item.Kind);
            foreach (RegulationModuleKind kind in source.Keys.Union(target.Keys).OrderBy(value => value))
            {
                RegulationPackageModuleManifest oldItem;
                RegulationPackageModuleManifest newItem;
                bool hasOld = source.TryGetValue(kind, out oldItem);
                bool hasNew = target.TryGetValue(kind, out newItem);
                if (!hasOld)
                {
                    changes.Add(new RegulationPackageChange(
                        RegulationPackageChangeKind.ModuleAdded,
                        RegulationPackageChangeImpact.CalculationData,
                        kind.ToString(),
                        string.Empty,
                        DescribeModule(newItem),
                        kind));
                }
                else if (!hasNew)
                {
                    changes.Add(new RegulationPackageChange(
                        RegulationPackageChangeKind.ModuleRemoved,
                        RegulationPackageChangeImpact.CalculationData,
                        kind.ToString(),
                        DescribeModule(oldItem),
                        string.Empty,
                        kind));
                }
                else
                {
                    string oldValue = DescribeModule(oldItem);
                    string newValue = DescribeModule(newItem);
                    if (!string.Equals(oldValue, newValue, StringComparison.Ordinal))
                    {
                        changes.Add(new RegulationPackageChange(
                            RegulationPackageChangeKind.ModuleChanged,
                            RegulationPackageChangeImpact.CalculationData,
                            kind.ToString(),
                            oldValue,
                            newValue,
                            kind));
                    }
                }
            }
        }

        private static void AddMetadataChange(
            List<RegulationPackageChange> changes,
            string key,
            string oldValue,
            string newValue)
        {
            if (string.Equals(oldValue ?? string.Empty, newValue ?? string.Empty, StringComparison.Ordinal))
                return;
            changes.Add(new RegulationPackageChange(
                RegulationPackageChangeKind.PackageMetadataChanged,
                RegulationPackageChangeImpact.Informational,
                key,
                oldValue,
                newValue,
                null));
        }

        private static string DescribeSource(RegulationPackageSourceDocument source)
        {
            return string.Join("|", new[]
            {
                source.Title,
                source.Publisher,
                FormatDate(source.IssuedDate),
                FormatDate(source.EffectiveFrom),
                FormatDate(source.EffectiveTo),
                source.OfficialUri,
                source.ContentChecksum
            });
        }

        private static string DescribeModule(RegulationPackageModuleManifest module)
        {
            return string.Join("|", new[]
            {
                module.ModuleId,
                module.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                module.DataVersion,
                module.RecordCount.ToString(CultureInfo.InvariantCulture),
                module.ContentChecksum
            });
        }

        private static string FormatDate(DateTime value)
        {
            return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        private static string FormatDate(DateTime? value)
        {
            return value.HasValue ? FormatDate(value.Value) : string.Empty;
        }

        private static void ValidatePackage(RegulationPackage package, string parameterName)
        {
            RegulationPackageValidationResult validation = RegulationPackageValidator.Validate(package);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), parameterName);
        }
    }

    public sealed class RegulationPackageMigrationPlan
    {
        private RegulationPackageMigrationPlan(
            string planId,
            RegulationPackage source,
            RegulationPackage target,
            RegulationPackageDiffResult diff)
        {
            PlanId = planId;
            Source = source;
            Target = target;
            Diff = diff;
        }

        public string PlanId { get; }
        public RegulationPackage Source { get; }
        public RegulationPackage Target { get; }
        public RegulationPackageDiffResult Diff { get; }

        public static RegulationPackageMigrationPlan Create(
            RegulationPackage source,
            RegulationPackage target)
        {
            RegulationPackageDiffResult diff = RegulationPackageDiffer.Compare(source, target);
            if (target.Status != RegulationPackageStatus.Published &&
                target.Status != RegulationPackageStatus.Superseded)
            {
                throw new ArgumentException("Target package phai la Published hoac Superseded.", nameof(target));
            }
            string fingerprint = string.Join("\n", new[]
            {
                source.PackageId,
                source.DataVersion,
                source.PackageChecksum,
                target.PackageId,
                target.DataVersion,
                target.PackageChecksum,
                diff.Changes.Count.ToString(CultureInfo.InvariantCulture)
            });
            string planId = "MIG-" + RegulationPackageSerializer.ComputeSha256(fingerprint).Substring(0, 24);
            return new RegulationPackageMigrationPlan(planId, source, target, diff);
        }
    }
}
