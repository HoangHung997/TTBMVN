using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum RegulationDataRecordChangeKind
    {
        Added,
        Removed,
        Changed
    }

    public sealed class RegulationDataFieldChange
    {
        public RegulationDataFieldChange(string field, string oldValue, string newValue)
        {
            Field = field ?? string.Empty;
            OldValue = oldValue ?? string.Empty;
            NewValue = newValue ?? string.Empty;
        }

        public string Field { get; }
        public string OldValue { get; }
        public string NewValue { get; }
    }

    public sealed class RegulationDataRecordChange
    {
        internal RegulationDataRecordChange(
            RegulationModuleKind moduleKind,
            string key,
            RegulationDataRecordChangeKind kind,
            IEnumerable<RegulationDataFieldChange> fieldChanges)
        {
            ModuleKind = moduleKind;
            Key = key ?? string.Empty;
            Kind = kind;
            FieldChanges = new ReadOnlyCollection<RegulationDataFieldChange>(
                (fieldChanges ?? Enumerable.Empty<RegulationDataFieldChange>()).ToList());
        }

        public RegulationModuleKind ModuleKind { get; }
        public string Key { get; }
        public RegulationDataRecordChangeKind Kind { get; }
        public IReadOnlyList<RegulationDataFieldChange> FieldChanges { get; }
    }

    public sealed class RegulationDataBundleDiffResult
    {
        internal RegulationDataBundleDiffResult(IEnumerable<RegulationDataRecordChange> changes)
        {
            Changes = new ReadOnlyCollection<RegulationDataRecordChange>(
                (changes ?? Enumerable.Empty<RegulationDataRecordChange>()).ToList());
        }

        public bool HasChanges => Changes.Count > 0;
        public IReadOnlyList<RegulationDataRecordChange> Changes { get; }
        public int AddedCount => Changes.Count(change => change.Kind == RegulationDataRecordChangeKind.Added);
        public int RemovedCount => Changes.Count(change => change.Kind == RegulationDataRecordChangeKind.Removed);
        public int ChangedCount => Changes.Count(change => change.Kind == RegulationDataRecordChangeKind.Changed);
    }

    public static class RegulationDataBundleDiffer
    {
        public static RegulationDataBundleDiffResult Compare(
            RegulationPackageBundle source,
            RegulationPackageBundle target)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (target == null)
                throw new ArgumentNullException(nameof(target));

            var changes = new List<RegulationDataRecordChange>();
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                RegulationDataModule oldModule;
                RegulationDataModule newModule;
                source.Modules.TryGetValue(kind, out oldModule);
                target.Modules.TryGetValue(kind, out newModule);
                CompareModule(kind, oldModule, newModule, changes);
            }
            return new RegulationDataBundleDiffResult(changes
                .OrderBy(change => change.ModuleKind)
                .ThenBy(change => change.Key, StringComparer.Ordinal));
        }

        private static void CompareModule(
            RegulationModuleKind kind,
            RegulationDataModule source,
            RegulationDataModule target,
            ICollection<RegulationDataRecordChange> changes)
        {
            Dictionary<string, RegulationDataRecord> oldRecords = ToDictionary(source);
            Dictionary<string, RegulationDataRecord> newRecords = ToDictionary(target);
            foreach (string key in oldRecords.Keys.Union(newRecords.Keys, StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal))
            {
                RegulationDataRecord oldRecord;
                RegulationDataRecord newRecord;
                bool hasOld = oldRecords.TryGetValue(key, out oldRecord);
                bool hasNew = newRecords.TryGetValue(key, out newRecord);
                if (!hasOld)
                {
                    changes.Add(new RegulationDataRecordChange(
                        kind,
                        key,
                        RegulationDataRecordChangeKind.Added,
                        DescribeFields(null, newRecord)));
                }
                else if (!hasNew)
                {
                    changes.Add(new RegulationDataRecordChange(
                        kind,
                        key,
                        RegulationDataRecordChangeKind.Removed,
                        DescribeFields(oldRecord, null)));
                }
                else
                {
                    List<RegulationDataFieldChange> fields = DescribeFields(oldRecord, newRecord).ToList();
                    if (fields.Count > 0)
                    {
                        changes.Add(new RegulationDataRecordChange(
                            kind,
                            key,
                            RegulationDataRecordChangeKind.Changed,
                            fields));
                    }
                }
            }
        }

        private static Dictionary<string, RegulationDataRecord> ToDictionary(RegulationDataModule module)
        {
            return (module?.Records ?? new RegulationDataRecord[0])
                .ToDictionary(record => record.Key, StringComparer.Ordinal);
        }

        private static IEnumerable<RegulationDataFieldChange> DescribeFields(
            RegulationDataRecord oldRecord,
            RegulationDataRecord newRecord)
        {
            var fields = new[]
            {
                Field("RecordType", oldRecord?.RecordType, newRecord?.RecordType),
                Field("Unit", oldRecord?.Unit, newRecord?.Unit),
                Field("Title", oldRecord?.Title, newRecord?.Title),
                Field("Data", oldRecord?.Data, newRecord?.Data),
                Field("SourceDocumentId", oldRecord?.Source?.DocumentId, newRecord?.Source?.DocumentId),
                Field(
                    "SourcePageFrom",
                    FormatInt(oldRecord?.Source?.PageFrom),
                    FormatInt(newRecord?.Source?.PageFrom)),
                Field(
                    "SourcePageTo",
                    FormatInt(oldRecord?.Source?.PageTo),
                    FormatInt(newRecord?.Source?.PageTo)),
                Field("SourceSection", oldRecord?.Source?.Section, newRecord?.Source?.Section),
                Field(
                    "Verification",
                    oldRecord == null ? string.Empty : oldRecord.Verification.ToString(),
                    newRecord == null ? string.Empty : newRecord.Verification.ToString())
            };
            return fields.Where(field => field != null);
        }

        private static RegulationDataFieldChange Field(string name, string oldValue, string newValue)
        {
            string normalizedOld = oldValue ?? string.Empty;
            string normalizedNew = newValue ?? string.Empty;
            return string.Equals(normalizedOld, normalizedNew, StringComparison.Ordinal)
                ? null
                : new RegulationDataFieldChange(name, normalizedOld, normalizedNew);
        }

        private static string FormatInt(int? value)
        {
            return value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : string.Empty;
        }
    }
}
