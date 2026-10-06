using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAddIn1.Core
{
    public enum ResultAuditKind
    {
        EstimateLine = 1,
        EstimateBlock = 2,
        CostComponent = 3,
        CostSummary = 4
    }

    public enum ResultAuditSourceKind
    {
        Regulation = 1,
        Price = 2,
        Quantity = 3,
        Output = 4,
        External = 5,
        Metadata = 6
    }

    public sealed class ResultAuditSource
    {
        public ResultAuditSource(
            ResultAuditSourceKind kind,
            string documentId,
            int pageFrom,
            int pageTo,
            string section,
            string reference)
        {
            Kind = kind;
            DocumentId = Normalize(documentId);
            PageFrom = pageFrom;
            PageTo = pageTo;
            Section = Normalize(section);
            Reference = Normalize(reference);
        }

        public ResultAuditSourceKind Kind { get; }
        public string DocumentId { get; }
        public int PageFrom { get; }
        public int PageTo { get; }
        public string Section { get; }
        public string Reference { get; }

        public string SortKey => string.Join("|", new[]
        {
            ((int)Kind).ToString("D2", CultureInfo.InvariantCulture),
            DocumentId,
            PageFrom.ToString("D6", CultureInfo.InvariantCulture),
            PageTo.ToString("D6", CultureInfo.InvariantCulture),
            Section,
            Reference
        });

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim();
        }
    }

    public sealed class ResultAuditEntry
    {
        public ResultAuditEntry(
            string entryId,
            string scopeId,
            ResultAuditKind kind,
            string worksheetCodeName,
            string worksheetRoleId,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            string label,
            string packageId,
            string packageVersion,
            string packageChecksum,
            string priceProfileId,
            string priceProfileVersion,
            string priceProfileChecksum,
            string normKey,
            string variantCode,
            IEnumerable<ResultAuditSource> sources)
        {
            EntryId = Normalize(entryId);
            ScopeId = Normalize(scopeId);
            Kind = kind;
            WorksheetCodeName = Normalize(worksheetCodeName);
            WorksheetRoleId = Normalize(worksheetRoleId);
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
            Label = Normalize(label);
            PackageId = Normalize(packageId);
            PackageVersion = Normalize(packageVersion);
            PackageChecksum = Normalize(packageChecksum).ToUpperInvariant();
            PriceProfileId = Normalize(priceProfileId);
            PriceProfileVersion = Normalize(priceProfileVersion);
            PriceProfileChecksum = Normalize(priceProfileChecksum).ToUpperInvariant();
            NormKey = Normalize(normKey);
            VariantCode = Normalize(variantCode);
            Sources = new ReadOnlyCollection<ResultAuditSource>(
                (sources ?? Enumerable.Empty<ResultAuditSource>())
                    .Where(source => source != null)
                    .OrderBy(source => source.SortKey, StringComparer.Ordinal)
                    .ToList());
        }

        public string EntryId { get; }
        public string ScopeId { get; }
        public ResultAuditKind Kind { get; }
        public string WorksheetCodeName { get; }
        public string WorksheetRoleId { get; }
        public int FirstRow { get; }
        public int FirstColumn { get; }
        public int LastRow { get; }
        public int LastColumn { get; }
        public string Label { get; }
        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public string PriceProfileId { get; }
        public string PriceProfileVersion { get; }
        public string PriceProfileChecksum { get; }
        public string NormKey { get; }
        public string VariantCode { get; }
        public IReadOnlyList<ResultAuditSource> Sources { get; }
        public long CellCount => checked((long)(LastRow - FirstRow + 1) * (LastColumn - FirstColumn + 1));

        public bool Contains(string worksheetCodeName, int row, int column)
        {
            return string.Equals(
                    WorksheetCodeName,
                    (worksheetCodeName ?? string.Empty).Trim(),
                    StringComparison.Ordinal) &&
                row >= FirstRow && row <= LastRow &&
                column >= FirstColumn && column <= LastColumn;
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim();
        }
    }

    public sealed class ResultAuditTrail
    {
        public const int CurrentSchemaVersion = 1;

        public ResultAuditTrail(IEnumerable<ResultAuditEntry> entries)
        {
            SchemaVersion = CurrentSchemaVersion;
            Entries = new ReadOnlyCollection<ResultAuditEntry>(
                (entries ?? Enumerable.Empty<ResultAuditEntry>())
                    .Where(entry => entry != null)
                    .OrderBy(entry => entry.EntryId, StringComparer.Ordinal)
                    .ToList());
        }

        public int SchemaVersion { get; }
        public IReadOnlyList<ResultAuditEntry> Entries { get; }

        public ResultAuditTrail ReplaceScope(string scopeId, IEnumerable<ResultAuditEntry> entries)
        {
            string scope = (scopeId ?? string.Empty).Trim();
            var combined = Entries
                .Where(entry => !string.Equals(entry.ScopeId, scope, StringComparison.Ordinal))
                .Concat(entries ?? Enumerable.Empty<ResultAuditEntry>());
            var result = new ResultAuditTrail(combined);
            ResultAuditTrailValidator.ValidateRequired(result);
            return result;
        }

        public ResultAuditEntry FindMostSpecific(string worksheetCodeName, int row, int column)
        {
            return Entries
                .Where(entry => entry.Contains(worksheetCodeName, row, column))
                .OrderBy(entry => entry.CellCount)
                .ThenBy(entry => entry.Kind)
                .ThenBy(entry => entry.EntryId, StringComparer.Ordinal)
                .FirstOrDefault();
        }
    }

    public sealed class ResultAuditTrailValidationResult
    {
        internal ResultAuditTrailValidationResult(IEnumerable<string> errors)
        {
            Errors = new ReadOnlyCollection<string>(errors.ToList());
        }

        public IReadOnlyList<string> Errors { get; }
        public bool IsValid => Errors.Count == 0;
    }

    public static class ResultAuditTrailValidator
    {
        public static ResultAuditTrailValidationResult Validate(ResultAuditTrail trail)
        {
            var errors = new List<string>();
            if (trail == null)
            {
                errors.Add("Result audit trail khong duoc null.");
                return new ResultAuditTrailValidationResult(errors);
            }
            if (trail.SchemaVersion != ResultAuditTrail.CurrentSchemaVersion)
                errors.Add("Result audit schema khong duoc ho tro.");

            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (ResultAuditEntry entry in trail.Entries)
            {
                if (entry.EntryId.Length == 0 || !ids.Add(entry.EntryId))
                    errors.Add("EntryId trong hoac bi trung: " + entry.EntryId + ".");
                if (entry.ScopeId.Length == 0 || entry.WorksheetCodeName.Length == 0 ||
                    entry.WorksheetRoleId.Length == 0 || entry.Label.Length == 0)
                    errors.Add("Audit entry thieu scope, worksheet role hoac label: " + entry.EntryId + ".");
                if (!Enum.IsDefined(typeof(ResultAuditKind), entry.Kind))
                    errors.Add("Audit kind khong hop le: " + entry.EntryId + ".");
                if (entry.FirstRow < 1 || entry.FirstColumn < 1 ||
                    entry.LastRow < entry.FirstRow || entry.LastColumn < entry.FirstColumn)
                    errors.Add("Vung audit khong hop le: " + entry.EntryId + ".");
                if (entry.PackageId.Length == 0 || entry.PackageVersion.Length == 0 ||
                    !RegulationPackageValidator.IsSha256(entry.PackageChecksum))
                    errors.Add("Danh tinh package audit khong hop le: " + entry.EntryId + ".");

                bool hasProfile = entry.PriceProfileId.Length > 0 ||
                    entry.PriceProfileVersion.Length > 0 || entry.PriceProfileChecksum.Length > 0;
                if (hasProfile && (entry.PriceProfileId.Length == 0 ||
                    entry.PriceProfileVersion.Length == 0 ||
                    !RegulationPackageValidator.IsSha256(entry.PriceProfileChecksum)))
                    errors.Add("Danh tinh PriceProfile audit khong hop le: " + entry.EntryId + ".");

                foreach (ResultAuditSource source in entry.Sources)
                {
                    if (!Enum.IsDefined(typeof(ResultAuditSourceKind), source.Kind) ||
                        source.PageFrom < 0 || source.PageTo < source.PageFrom ||
                        (source.DocumentId.Length == 0 && source.Reference.Length == 0))
                        errors.Add("Nguon audit khong hop le: " + entry.EntryId + ".");
                }
            }
            return new ResultAuditTrailValidationResult(errors);
        }

        internal static void ValidateRequired(ResultAuditTrail trail)
        {
            ResultAuditTrailValidationResult result = Validate(trail);
            if (!result.IsValid)
                throw new ArgumentException(string.Join(" ", result.Errors), nameof(trail));
        }
    }

    public static class ResultAuditTrailSerializer
    {
        private const string Magic = "TTBMVN_RESULT_AUDIT";

        public static string Serialize(ResultAuditTrail trail)
        {
            ResultAuditTrailValidator.ValidateRequired(trail);
            var lines = new List<string>
            {
                Magic,
                "schema=" + trail.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                "entries=" + trail.Entries.Count.ToString(CultureInfo.InvariantCulture)
            };
            foreach (ResultAuditEntry entry in trail.Entries)
            {
                lines.Add(string.Join("|", new[]
                {
                    "entry", Encode(entry.EntryId), Encode(entry.ScopeId),
                    ((int)entry.Kind).ToString(CultureInfo.InvariantCulture),
                    Encode(entry.WorksheetCodeName), Encode(entry.WorksheetRoleId),
                    entry.FirstRow.ToString(CultureInfo.InvariantCulture),
                    entry.FirstColumn.ToString(CultureInfo.InvariantCulture),
                    entry.LastRow.ToString(CultureInfo.InvariantCulture),
                    entry.LastColumn.ToString(CultureInfo.InvariantCulture),
                    Encode(entry.Label), Encode(entry.PackageId), Encode(entry.PackageVersion),
                    entry.PackageChecksum, Encode(entry.PriceProfileId),
                    Encode(entry.PriceProfileVersion), entry.PriceProfileChecksum,
                    Encode(entry.NormKey), Encode(entry.VariantCode)
                }));
                foreach (ResultAuditSource source in entry.Sources)
                {
                    lines.Add(string.Join("|", new[]
                    {
                        "source", Encode(entry.EntryId),
                        ((int)source.Kind).ToString(CultureInfo.InvariantCulture),
                        Encode(source.DocumentId),
                        source.PageFrom.ToString(CultureInfo.InvariantCulture),
                        source.PageTo.ToString(CultureInfo.InvariantCulture),
                        Encode(source.Section), Encode(source.Reference)
                    }));
                }
            }
            return string.Join("\n", lines);
        }

        public static ResultAuditTrail Deserialize(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
                throw new InvalidDataException("Result audit payload trong.");
            string[] lines = payload.Replace("\r\n", "\n").Replace('\r', '\n')
                .Split(new[] { '\n' }, StringSplitOptions.None);
            if (lines.Length < 3 || !string.Equals(lines[0], Magic, StringComparison.Ordinal))
                throw new InvalidDataException("Result audit payload khong dung dinh dang.");
            int schema = ParseHeader(lines[1], "schema");
            int expectedCount = ParseHeader(lines[2], "entries");
            if (schema != ResultAuditTrail.CurrentSchemaVersion || expectedCount < 0)
                throw new InvalidDataException("Result audit header khong hop le.");

            var builders = new Dictionary<string, EntryBuilder>(StringComparer.Ordinal);
            var order = new List<string>();
            for (int index = 3; index < lines.Length; index++)
            {
                string[] parts = lines[index].Split('|');
                if (parts.Length == 19 && string.Equals(parts[0], "entry", StringComparison.Ordinal))
                {
                    string id = Decode(parts[1]);
                    if (builders.ContainsKey(id))
                        throw new InvalidDataException("Result audit trung EntryId: " + id + ".");
                    builders.Add(id, new EntryBuilder(parts));
                    order.Add(id);
                }
                else if (parts.Length == 8 && string.Equals(parts[0], "source", StringComparison.Ordinal))
                {
                    string id = Decode(parts[1]);
                    EntryBuilder builder;
                    if (!builders.TryGetValue(id, out builder))
                        throw new InvalidDataException("Result audit source khong co entry: " + id + ".");
                    builder.Sources.Add(new ResultAuditSource(
                        ParseEnum<ResultAuditSourceKind>(parts[2]),
                        Decode(parts[3]), ParseInt(parts[4]), ParseInt(parts[5]),
                        Decode(parts[6]), Decode(parts[7])));
                }
                else
                {
                    throw new InvalidDataException("Result audit co dong khong hop le tai " + (index + 1) + ".");
                }
            }
            if (builders.Count != expectedCount)
                throw new InvalidDataException("Result audit entry count khong khop.");

            var trail = new ResultAuditTrail(order.Select(id => builders[id].Build()));
            ResultAuditTrailValidationResult validation = ResultAuditTrailValidator.Validate(trail);
            if (!validation.IsValid)
                throw new InvalidDataException(string.Join(" ", validation.Errors));
            return trail;
        }

        public static string ComputeChecksum(string payload)
        {
            if (payload == null)
                throw new ArgumentNullException(nameof(payload));
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(payload));
                var output = new StringBuilder(hash.Length * 2);
                foreach (byte value in hash)
                    output.Append(value.ToString("X2", CultureInfo.InvariantCulture));
                return output.ToString();
            }
        }

        private static int ParseHeader(string line, string name)
        {
            string prefix = name + "=";
            if (line == null || !line.StartsWith(prefix, StringComparison.Ordinal))
                throw new InvalidDataException("Result audit thieu header " + name + ".");
            return ParseInt(line.Substring(prefix.Length));
        }

        private static int ParseInt(string value)
        {
            int result;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
                throw new InvalidDataException("Result audit integer khong hop le.");
            return result;
        }

        private static T ParseEnum<T>(string value) where T : struct
        {
            int number = ParseInt(value);
            T result = (T)Enum.ToObject(typeof(T), number);
            if (!Enum.IsDefined(typeof(T), result))
                throw new InvalidDataException("Result audit enum khong hop le.");
            return result;
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value ?? string.Empty));
            }
            catch (FormatException ex)
            {
                throw new InvalidDataException("Result audit Base64 khong hop le.", ex);
            }
        }

        private sealed class EntryBuilder
        {
            private readonly string[] parts;

            public EntryBuilder(string[] parts)
            {
                this.parts = parts;
            }

            public List<ResultAuditSource> Sources { get; } = new List<ResultAuditSource>();

            public ResultAuditEntry Build()
            {
                return new ResultAuditEntry(
                    Decode(parts[1]), Decode(parts[2]), ParseEnum<ResultAuditKind>(parts[3]),
                    Decode(parts[4]), Decode(parts[5]), ParseInt(parts[6]), ParseInt(parts[7]),
                    ParseInt(parts[8]), ParseInt(parts[9]), Decode(parts[10]), Decode(parts[11]),
                    Decode(parts[12]), parts[13], Decode(parts[14]), Decode(parts[15]),
                    parts[16], Decode(parts[17]), Decode(parts[18]), Sources);
            }
        }
    }
}
