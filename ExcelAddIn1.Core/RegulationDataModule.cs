using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAddIn1.Core
{
    public enum RegulationDataVerification
    {
        Unverified = 0,
        VerifiedAgainstOfficialSource = 1,
        DerivedFromVerifiedRecords = 2
    }

    public sealed class RegulationSourceLocator
    {
        public RegulationSourceLocator(
            string documentId,
            int pageFrom,
            int pageTo,
            string section)
        {
            DocumentId = documentId;
            PageFrom = pageFrom;
            PageTo = pageTo;
            Section = section;
        }

        public string DocumentId { get; }
        public int PageFrom { get; }
        public int PageTo { get; }
        public string Section { get; }
    }

    public sealed class RegulationDataRecord
    {
        public RegulationDataRecord(
            string key,
            string recordType,
            string unit,
            string title,
            string data,
            RegulationSourceLocator source,
            RegulationDataVerification verification)
        {
            Key = key;
            RecordType = recordType;
            Unit = unit;
            Title = title;
            Data = data;
            Source = source;
            Verification = verification;
        }

        public string Key { get; }
        public string RecordType { get; }
        public string Unit { get; }
        public string Title { get; }
        public string Data { get; }
        public RegulationSourceLocator Source { get; }
        public RegulationDataVerification Verification { get; }
    }

    public sealed class RegulationDataModule
    {
        public const int CurrentSchemaVersion = 1;

        private RegulationDataModule(
            int schemaVersion,
            RegulationModuleKind kind,
            string dataVersion,
            IEnumerable<RegulationDataRecord> records,
            string contentChecksum)
        {
            SchemaVersion = schemaVersion;
            Kind = kind;
            DataVersion = dataVersion;
            Records = new ReadOnlyCollection<RegulationDataRecord>(records.ToList());
            ContentChecksum = contentChecksum;
        }

        public int SchemaVersion { get; }
        public RegulationModuleKind Kind { get; }
        public string DataVersion { get; }
        public IReadOnlyList<RegulationDataRecord> Records { get; }
        public string ContentChecksum { get; }

        public static RegulationDataModule Create(
            RegulationModuleKind kind,
            string dataVersion,
            IEnumerable<RegulationDataRecord> records)
        {
            RegulationDataRecord[] normalizedRecords = (records ?? Enumerable.Empty<RegulationDataRecord>())
                .Select(NormalizeRecord)
                .OrderBy(record => record.Key, StringComparer.Ordinal)
                .ToArray();
            var unsigned = new RegulationDataModule(
                CurrentSchemaVersion,
                kind,
                (dataVersion ?? string.Empty).Trim(),
                normalizedRecords,
                string.Empty);
            RegulationDataValidationResult validation = RegulationDataValidator.ValidateForSealing(unsigned);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(records));
            return new RegulationDataModule(
                unsigned.SchemaVersion,
                unsigned.Kind,
                unsigned.DataVersion,
                unsigned.Records,
                RegulationDataModuleSerializer.ComputeChecksum(unsigned));
        }

        internal static RegulationDataModule Rehydrate(
            int schemaVersion,
            RegulationModuleKind kind,
            string dataVersion,
            IEnumerable<RegulationDataRecord> records,
            string contentChecksum)
        {
            return new RegulationDataModule(
                schemaVersion,
                kind,
                dataVersion,
                records ?? Enumerable.Empty<RegulationDataRecord>(),
                contentChecksum);
        }

        private static RegulationDataRecord NormalizeRecord(RegulationDataRecord record)
        {
            if (record == null)
                return null;
            RegulationSourceLocator source = record.Source == null
                ? null
                : new RegulationSourceLocator(
                    (record.Source.DocumentId ?? string.Empty).Trim(),
                    record.Source.PageFrom,
                    record.Source.PageTo,
                    (record.Source.Section ?? string.Empty).Trim());
            return new RegulationDataRecord(
                (record.Key ?? string.Empty).Trim(),
                (record.RecordType ?? string.Empty).Trim(),
                (record.Unit ?? string.Empty).Trim(),
                (record.Title ?? string.Empty).Trim(),
                (record.Data ?? string.Empty).Trim(),
                source,
                record.Verification);
        }
    }

    public sealed class RegulationDataValidationResult
    {
        internal RegulationDataValidationResult(IEnumerable<string> errors)
        {
            Errors = new ReadOnlyCollection<string>(errors.ToList());
        }

        public bool IsValid => Errors.Count == 0;
        public IReadOnlyList<string> Errors { get; }
    }

    public static class RegulationDataValidator
    {
        private static readonly Regex IdPattern = new Regex(
            "^[A-Za-z0-9][A-Za-z0-9._-]{1,159}$",
            RegexOptions.CultureInvariant);
        private static readonly Regex VersionPattern = new Regex(
            "^[0-9]+\\.[0-9]+\\.[0-9]+(?:-[0-9A-Za-z.-]+)?$",
            RegexOptions.CultureInvariant);
        private static readonly Regex Sha256Pattern = new Regex(
            "^[0-9A-Fa-f]{64}$",
            RegexOptions.CultureInvariant);

        public static RegulationDataValidationResult Validate(RegulationDataModule module)
        {
            return ValidateInternal(module, true, null, false);
        }

        public static RegulationDataValidationResult ValidateAgainstPackage(
            RegulationDataModule module,
            RegulationPackage package,
            bool requireOfficialVerification)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));
            var sourceIds = new HashSet<string>(
                package.Sources.Select(source => source.DocumentId),
                StringComparer.Ordinal);
            RegulationDataValidationResult basic = ValidateInternal(
                module,
                true,
                sourceIds,
                requireOfficialVerification);
            var errors = basic.Errors.ToList();
            RegulationPackageModuleManifest manifest = package.Modules.FirstOrDefault(
                candidate => candidate.Kind == module?.Kind);
            if (manifest == null)
            {
                errors.Add("Package thieu manifest cho module data.");
            }
            else if (module != null)
            {
                if (manifest.RecordCount != module.Records.Count)
                    errors.Add("RecordCount module khong khop manifest package.");
                if (!string.Equals(
                    manifest.ContentChecksum,
                    RegulationDataModuleSerializer.ComputeSerializedChecksum(module),
                    StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add("Checksum module khong khop manifest package.");
                }
            }
            return new RegulationDataValidationResult(errors);
        }

        internal static RegulationDataValidationResult ValidateForSealing(RegulationDataModule module)
        {
            return ValidateInternal(module, false, null, false);
        }

        private static RegulationDataValidationResult ValidateInternal(
            RegulationDataModule module,
            bool requireChecksum,
            ISet<string> packageSourceIds,
            bool requireOfficialVerification)
        {
            var errors = new List<string>();
            if (module == null)
            {
                errors.Add("Regulation data module null.");
                return new RegulationDataValidationResult(errors);
            }
            if (module.SchemaVersion != RegulationDataModule.CurrentSchemaVersion)
                errors.Add("Schema data module khong duoc ho tro.");
            if (!Enum.IsDefined(typeof(RegulationModuleKind), module.Kind))
                errors.Add("Module kind khong hop le.");
            if (!VersionPattern.IsMatch(module.DataVersion ?? string.Empty))
                errors.Add("DataVersion khong hop le.");
            if (module.Records == null || module.Records.Count == 0)
                errors.Add("Data module khong co record.");
            if (module.Records != null && module.Records.Count > 100000)
                errors.Add("Data module vuot gioi han record.");

            var keys = new HashSet<string>(StringComparer.Ordinal);
            if (module.Records != null)
            {
                for (int index = 0; index < module.Records.Count; index++)
                    ValidateRecord(
                        module.Records[index],
                        index,
                        keys,
                        packageSourceIds,
                        requireOfficialVerification,
                        errors);
            }

            if (requireChecksum)
            {
                if (!Sha256Pattern.IsMatch(module.ContentChecksum ?? string.Empty))
                    errors.Add("ContentChecksum data module khong hop le.");
                else if (!string.Equals(
                    module.ContentChecksum,
                    RegulationDataModuleSerializer.ComputeChecksum(module),
                    StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add("ContentChecksum data module khong khop noi dung.");
                }
            }
            return new RegulationDataValidationResult(errors);
        }

        private static void ValidateRecord(
            RegulationDataRecord record,
            int index,
            ISet<string> keys,
            ISet<string> packageSourceIds,
            bool requireOfficialVerification,
            ICollection<string> errors)
        {
            string prefix = "Record " + index.ToString(CultureInfo.InvariantCulture) + ": ";
            if (record == null)
            {
                errors.Add(prefix + "null.");
                return;
            }
            if (!IdPattern.IsMatch(record.Key ?? string.Empty))
                errors.Add(prefix + "key khong hop le.");
            else if (!keys.Add(record.Key))
                errors.Add(prefix + "key bi trung.");
            if (!IdPattern.IsMatch(record.RecordType ?? string.Empty))
                errors.Add(prefix + "record type khong hop le.");
            if (!string.IsNullOrEmpty(record.Unit) && !IdPattern.IsMatch(record.Unit))
                errors.Add(prefix + "unit khong hop le.");
            if (string.IsNullOrWhiteSpace(record.Title) || record.Title.Length > 1000)
                errors.Add(prefix + "title khong hop le.");
            if (record.Data == null || record.Data.Length > 1000000)
                errors.Add(prefix + "data khong hop le.");
            if (record.Source == null)
            {
                errors.Add(prefix + "source locator null.");
            }
            else
            {
                if (!IdPattern.IsMatch(record.Source.DocumentId ?? string.Empty))
                    errors.Add(prefix + "source document ID khong hop le.");
                if (record.Source.PageFrom < 1 || record.Source.PageTo < record.Source.PageFrom)
                    errors.Add(prefix + "source page khong hop le.");
                if (string.IsNullOrWhiteSpace(record.Source.Section) || record.Source.Section.Length > 1000)
                    errors.Add(prefix + "source section khong hop le.");
                if (packageSourceIds != null && !packageSourceIds.Contains(record.Source.DocumentId))
                    errors.Add(prefix + "source document khong thuoc package.");
            }
            if (!Enum.IsDefined(typeof(RegulationDataVerification), record.Verification) ||
                record.Verification == RegulationDataVerification.Unverified)
            {
                errors.Add(prefix + "record chua xac minh.");
            }
            if (requireOfficialVerification &&
                record.Verification != RegulationDataVerification.VerifiedAgainstOfficialSource)
            {
                errors.Add(prefix + "record khong duoc doi chieu truc tiep nguon chinh thuc.");
            }
        }
    }

    public static class RegulationDataModuleSerializer
    {
        private const string Magic = "TTBMVN_REGULATION_DATA_MODULE";

        public static string Serialize(RegulationDataModule module)
        {
            RegulationDataValidationResult validation = RegulationDataValidator.Validate(module);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(module));
            return BuildBody(module) + "\nchecksum=" + module.ContentChecksum.ToUpperInvariant();
        }

        public static RegulationDataModule Deserialize(string payload)
        {
            Dictionary<string, string> fields = Parse(payload);
            int schema = ParseInt(Get(fields, "schema"), "schema", 1, int.MaxValue);
            if (schema != RegulationDataModule.CurrentSchemaVersion)
                throw new InvalidDataException("Data module schema khong duoc ho tro.");
            int count = ParseInt(Get(fields, "recordCount"), "recordCount", 1, 100000);
            int kindValue = ParseInt(Get(fields, "kind"), "kind", int.MinValue, int.MaxValue);
            var records = new List<RegulationDataRecord>(count);
            for (int index = 0; index < count; index++)
            {
                string prefix = "record." + index.ToString(CultureInfo.InvariantCulture) + ".";
                records.Add(new RegulationDataRecord(
                    Decode(Get(fields, prefix + "key"), prefix + "key"),
                    Decode(Get(fields, prefix + "type"), prefix + "type"),
                    Decode(Get(fields, prefix + "unit"), prefix + "unit"),
                    Decode(Get(fields, prefix + "title"), prefix + "title"),
                    Decode(Get(fields, prefix + "data"), prefix + "data"),
                    new RegulationSourceLocator(
                        Decode(Get(fields, prefix + "sourceDocument"), prefix + "sourceDocument"),
                        ParseInt(Get(fields, prefix + "pageFrom"), prefix + "pageFrom", 1, int.MaxValue),
                        ParseInt(Get(fields, prefix + "pageTo"), prefix + "pageTo", 1, int.MaxValue),
                        Decode(Get(fields, prefix + "section"), prefix + "section")),
                    (RegulationDataVerification)ParseInt(
                        Get(fields, prefix + "verification"),
                        prefix + "verification",
                        int.MinValue,
                        int.MaxValue)));
            }
            int expectedFields = 5 + (count * 10);
            if (fields.Count != expectedFields)
                throw new InvalidDataException("Data module co field thua hoac count khong khop.");
            RegulationDataModule module = RegulationDataModule.Rehydrate(
                schema,
                (RegulationModuleKind)kindValue,
                Decode(Get(fields, "dataVersion"), "dataVersion"),
                records,
                Get(fields, "checksum"));
            RegulationDataValidationResult validation = RegulationDataValidator.Validate(module);
            if (!validation.IsValid)
                throw new InvalidDataException(string.Join(" ", validation.Errors));
            return module;
        }

        public static string ComputeChecksum(RegulationDataModule module)
        {
            if (module == null)
                throw new ArgumentNullException(nameof(module));
            return RegulationPackageSerializer.ComputeSha256(BuildBody(module));
        }

        public static string ComputeSerializedChecksum(RegulationDataModule module)
        {
            return RegulationPackageSerializer.ComputeSha256(Serialize(module));
        }

        private static string BuildBody(RegulationDataModule module)
        {
            RegulationDataRecord[] records = module.Records
                .OrderBy(record => record?.Key ?? string.Empty, StringComparer.Ordinal)
                .ToArray();
            var lines = new List<string>
            {
                Magic,
                "schema=" + module.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                "kind=" + ((int)module.Kind).ToString(CultureInfo.InvariantCulture),
                "dataVersion=" + Encode(module.DataVersion?.Trim() ?? string.Empty),
                "recordCount=" + records.Length.ToString(CultureInfo.InvariantCulture)
            };
            for (int index = 0; index < records.Length; index++)
            {
                RegulationDataRecord record = records[index];
                string prefix = "record." + index.ToString(CultureInfo.InvariantCulture) + ".";
                lines.Add(prefix + "key=" + Encode(record?.Key?.Trim() ?? string.Empty));
                lines.Add(prefix + "type=" + Encode(record?.RecordType?.Trim() ?? string.Empty));
                lines.Add(prefix + "unit=" + Encode(record?.Unit?.Trim() ?? string.Empty));
                lines.Add(prefix + "title=" + Encode(record?.Title?.Trim() ?? string.Empty));
                lines.Add(prefix + "data=" + Encode(record?.Data?.Trim() ?? string.Empty));
                lines.Add(prefix + "sourceDocument=" + Encode(record?.Source?.DocumentId?.Trim() ?? string.Empty));
                lines.Add(prefix + "pageFrom=" + (record?.Source?.PageFrom ?? 0).ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "pageTo=" + (record?.Source?.PageTo ?? 0).ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "section=" + Encode(record?.Source?.Section?.Trim() ?? string.Empty));
                lines.Add(prefix + "verification=" + (record == null
                    ? string.Empty
                    : ((int)record.Verification).ToString(CultureInfo.InvariantCulture)));
            }
            return string.Join("\n", lines);
        }

        private static Dictionary<string, string> Parse(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
                throw new InvalidDataException("Data module payload trong.");
            string[] lines = payload.Replace("\r\n", "\n").Replace('\r', '\n')
                .Split(new[] { '\n' }, StringSplitOptions.None);
            if (lines.Length == 0 || !string.Equals(lines[0], Magic, StringComparison.Ordinal))
                throw new InvalidDataException("Data module payload khong dung dinh dang.");
            var fields = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int index = 1; index < lines.Length; index++)
            {
                int separator = lines[index].IndexOf('=');
                if (separator <= 0)
                    throw new InvalidDataException("Data module co dong khong hop le.");
                string key = lines[index].Substring(0, separator);
                if (fields.ContainsKey(key))
                    throw new InvalidDataException("Data module trung field: " + key + ".");
                fields.Add(key, lines[index].Substring(separator + 1));
            }
            return fields;
        }

        private static string Get(IDictionary<string, string> fields, string key)
        {
            string value;
            if (!fields.TryGetValue(key, out value))
                throw new InvalidDataException("Data module thieu field: " + key + ".");
            return value;
        }

        private static int ParseInt(string value, string field, int minimum, int maximum)
        {
            int parsed;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ||
                parsed < minimum || parsed > maximum)
            {
                throw new InvalidDataException("Data module field " + field + " khong hop le.");
            }
            return parsed;
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value, string field)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value ?? string.Empty));
            }
            catch (FormatException ex)
            {
                throw new InvalidDataException("Data module field " + field + " khong phai Base64.", ex);
            }
        }
    }
}
