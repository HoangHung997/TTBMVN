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
    public enum EstimateWorkEnvironment
    {
        Land = 1,
        Water = 2
    }

    public sealed class EstimateColumnMap
    {
        public EstimateColumnMap(
            int codeColumn,
            int descriptionColumn,
            int unitColumn,
            int quantityColumn,
            int acceptedQuantityColumn,
            int materialRateColumn,
            int laborRateColumn,
            int machineRateColumn,
            int materialAmountColumn,
            int laborAmountColumn,
            int machineAmountColumn)
        {
            CodeColumn = codeColumn;
            DescriptionColumn = descriptionColumn;
            UnitColumn = unitColumn;
            QuantityColumn = quantityColumn;
            AcceptedQuantityColumn = acceptedQuantityColumn;
            MaterialRateColumn = materialRateColumn;
            LaborRateColumn = laborRateColumn;
            MachineRateColumn = machineRateColumn;
            MaterialAmountColumn = materialAmountColumn;
            LaborAmountColumn = laborAmountColumn;
            MachineAmountColumn = machineAmountColumn;
        }

        public int CodeColumn { get; }
        public int DescriptionColumn { get; }
        public int UnitColumn { get; }
        public int QuantityColumn { get; }
        public int AcceptedQuantityColumn { get; }
        public int MaterialRateColumn { get; }
        public int LaborRateColumn { get; }
        public int MachineRateColumn { get; }
        public int MaterialAmountColumn { get; }
        public int LaborAmountColumn { get; }
        public int MachineAmountColumn { get; }

        public IEnumerable<int> AllColumns => new[]
        {
            CodeColumn,
            DescriptionColumn,
            UnitColumn,
            QuantityColumn,
            AcceptedQuantityColumn,
            MaterialRateColumn,
            LaborRateColumn,
            MachineRateColumn,
            MaterialAmountColumn,
            LaborAmountColumn,
            MachineAmountColumn
        };
    }

    public sealed class EstimateSourceBinding
    {
        public EstimateSourceBinding(
            string worksheetCodeName,
            string worksheetName,
            string sourceAddress,
            int firstDataRow,
            int lastDataRow,
            EstimateColumnMap columns)
        {
            WorksheetCodeName = (worksheetCodeName ?? string.Empty).Trim();
            WorksheetName = (worksheetName ?? string.Empty).Trim();
            SourceAddress = (sourceAddress ?? string.Empty).Trim();
            FirstDataRow = firstDataRow;
            LastDataRow = lastDataRow;
            Columns = columns ?? throw new ArgumentNullException(nameof(columns));
        }

        public string WorksheetCodeName { get; }
        public string WorksheetName { get; }
        public string SourceAddress { get; }
        public int FirstDataRow { get; }
        public int LastDataRow { get; }
        public EstimateColumnMap Columns { get; }
    }

    public sealed class EstimateWorkspaceRow
    {
        public EstimateWorkspaceRow(
            string rowId,
            string bindingName,
            int sourceRowHint,
            string workCode,
            string description,
            string unit,
            string quantityFormulaLocal,
            decimal quantity,
            string acceptedQuantityFormulaLocal,
            decimal acceptedQuantity,
            EstimateWorkEnvironment environment,
            MachineRateAudience laborAudience,
            string normKey,
            string variantCode,
            IEnumerable<string> conditions = null,
            IEnumerable<UnitRateResourceBinding> bindings = null)
        {
            RowId = (rowId ?? string.Empty).Trim();
            BindingName = (bindingName ?? string.Empty).Trim();
            SourceRowHint = sourceRowHint;
            WorkCode = (workCode ?? string.Empty).Trim();
            Description = (description ?? string.Empty).Trim();
            Unit = (unit ?? string.Empty).Trim();
            QuantityFormulaLocal = (quantityFormulaLocal ?? string.Empty).Trim();
            Quantity = quantity;
            AcceptedQuantityFormulaLocal = (acceptedQuantityFormulaLocal ?? string.Empty).Trim();
            AcceptedQuantity = acceptedQuantity;
            Environment = environment;
            LaborAudience = laborAudience;
            NormKey = (normKey ?? string.Empty).Trim();
            VariantCode = (variantCode ?? string.Empty).Trim();
            Conditions = new ReadOnlyCollection<string>(
                (conditions ?? Enumerable.Empty<string>())
                    .Select(value => (value ?? string.Empty).Trim())
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToList());
            Bindings = new ReadOnlyCollection<UnitRateResourceBinding>(
                (bindings ?? Enumerable.Empty<UnitRateResourceBinding>())
                    .Where(value => value != null)
                    .OrderBy(value => value.ResourceCode, StringComparer.Ordinal)
                    .ThenBy(value => value.PriceCode, StringComparer.Ordinal)
                    .ToList());
        }

        public string RowId { get; }
        public string BindingName { get; }
        public int SourceRowHint { get; }
        public string WorkCode { get; }
        public string Description { get; }
        public string Unit { get; }
        public string QuantityFormulaLocal { get; }
        public decimal Quantity { get; }
        public string AcceptedQuantityFormulaLocal { get; }
        public decimal AcceptedQuantity { get; }
        public EstimateWorkEnvironment Environment { get; }
        public MachineRateAudience LaborAudience { get; }
        public string NormKey { get; }
        public string VariantCode { get; }
        public IReadOnlyList<string> Conditions { get; }
        public IReadOnlyList<UnitRateResourceBinding> Bindings { get; }
        public bool IsTextRow => NormKey.Length == 0 && VariantCode.Length == 0;
        public bool HasCompleteNorm => NormKey.Length > 0 && VariantCode.Length > 0;
    }

    public sealed class EstimateWorkspace
    {
        public const int CurrentSchemaVersion = 1;

        public EstimateWorkspace(
            EstimateSourceBinding source,
            IEnumerable<EstimateWorkspaceRow> rows,
            DateTime updatedAtUtc)
        {
            SchemaVersion = CurrentSchemaVersion;
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Rows = new ReadOnlyCollection<EstimateWorkspaceRow>(
                (rows ?? Enumerable.Empty<EstimateWorkspaceRow>())
                    .Where(row => row != null)
                    .OrderBy(row => row.SourceRowHint)
                    .ToList());
            UpdatedAtUtc = NormalizeUtc(updatedAtUtc);
            EstimateWorkspaceValidator.ValidateRequired(this);
        }

        public int SchemaVersion { get; }
        public EstimateSourceBinding Source { get; }
        public IReadOnlyList<EstimateWorkspaceRow> Rows { get; }
        public DateTime UpdatedAtUtc { get; }

        private static DateTime NormalizeUtc(DateTime value)
        {
            if (value.Kind == DateTimeKind.Utc)
                return value;
            if (value.Kind == DateTimeKind.Local)
                return value.ToUniversalTime();
            return DateTime.SpecifyKind(value, DateTimeKind.Utc);
        }
    }

    public sealed class EstimateWorkspaceValidationResult
    {
        internal EstimateWorkspaceValidationResult(IEnumerable<string> errors)
        {
            Errors = new ReadOnlyCollection<string>(errors.ToList());
        }

        public IReadOnlyList<string> Errors { get; }
        public bool IsValid => Errors.Count == 0;
    }

    public static class EstimateWorkspaceValidator
    {
        public static EstimateWorkspaceValidationResult Validate(EstimateWorkspace workspace)
        {
            var errors = new List<string>();
            if (workspace == null)
            {
                errors.Add("EstimateWorkspace khong duoc null.");
                return new EstimateWorkspaceValidationResult(errors);
            }

            EstimateSourceBinding source = workspace.Source;
            if (source == null)
            {
                errors.Add("EstimateWorkspace thieu vung nguon.");
                return new EstimateWorkspaceValidationResult(errors);
            }
            if (source.WorksheetCodeName.Length == 0 || source.WorksheetName.Length == 0 ||
                source.SourceAddress.Length == 0)
            {
                errors.Add("Vung nguon phai co CodeName, ten sheet va dia chi.");
            }
            if (source.FirstDataRow < 1 || source.LastDataRow < source.FirstDataRow)
                errors.Add("Khoang dong nguon khong hop le.");
            if (source.Columns.DescriptionColumn < 1 || source.Columns.QuantityColumn < 1)
                errors.Add("Phai anh xa cot noi dung va khoi luong.");
            foreach (int column in source.Columns.AllColumns)
            {
                if (column < 0 || column > 16384)
                    errors.Add("So cot Excel khong hop le: " + column + ".");
            }

            var rowIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var bindingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (EstimateWorkspaceRow row in workspace.Rows)
            {
                if (row.RowId.Length == 0 || !rowIds.Add(row.RowId))
                    errors.Add("RowId trong hoac bi trung: " + row.RowId + ".");
                if (row.BindingName.Length == 0 || !bindingNames.Add(row.BindingName))
                    errors.Add("BindingName trong hoac bi trung: " + row.BindingName + ".");
                if (row.SourceRowHint < source.FirstDataRow || row.SourceRowHint > source.LastDataRow)
                    errors.Add(row.RowId + ": dong nguon nam ngoai vung quet.");
                if (row.WorkCode.Length == 0 && row.Description.Length == 0 && row.Unit.Length == 0 &&
                    row.QuantityFormulaLocal.Length == 0 && row.AcceptedQuantityFormulaLocal.Length == 0)
                {
                    errors.Add(row.RowId + ": dong nguon khong co du lieu.");
                }
                if (row.Quantity < 0m || row.AcceptedQuantity < 0m)
                    errors.Add(row.RowId + ": khoi luong khong duoc am.");
                if (!Enum.IsDefined(typeof(EstimateWorkEnvironment), row.Environment))
                    errors.Add(row.RowId + ": moi truong khong hop le.");
                if (!Enum.IsDefined(typeof(MachineRateAudience), row.LaborAudience))
                    errors.Add(row.RowId + ": doi tuong luong khong hop le.");
                if ((row.NormKey.Length == 0) != (row.VariantCode.Length == 0))
                    errors.Add(row.RowId + ": phai chon dong thoi ma dinh muc va ma chi tiet.");
                var bindingCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (UnitRateResourceBinding binding in row.Bindings)
                {
                    if (binding.ResourceCode.Length == 0 || binding.PriceCode.Length == 0 ||
                        binding.Reason.Length == 0 || binding.Reason.Length > 1024)
                    {
                        errors.Add(row.RowId + ": binding phai co ma hao phi, ma gia va ly do.");
                    }
                    if (!bindingCodes.Add(binding.ResourceCode))
                        errors.Add(row.RowId + ": binding bi trung cho " + binding.ResourceCode + ".");
                }
            }
            return new EstimateWorkspaceValidationResult(errors.Distinct().ToList());
        }

        public static void ValidateRequired(EstimateWorkspace workspace)
        {
            EstimateWorkspaceValidationResult validation = Validate(workspace);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(workspace));
        }
    }

    public sealed class EstimateRateIdentity : IEquatable<EstimateRateIdentity>
    {
        public EstimateRateIdentity(
            string packageId,
            string packageVersion,
            string packageChecksum,
            string normKey,
            string variantCode,
            EstimateWorkEnvironment environment,
            MachineRateAudience laborAudience,
            PriceProfile priceProfile,
            IEnumerable<string> conditions,
            IEnumerable<UnitRateResourceBinding> bindings = null)
        {
            if (priceProfile == null)
                throw new ArgumentNullException(nameof(priceProfile));
            PackageId = Required(packageId, nameof(packageId));
            PackageVersion = Required(packageVersion, nameof(packageVersion));
            PackageChecksum = Required(packageChecksum, nameof(packageChecksum));
            NormKey = Required(normKey, nameof(normKey));
            VariantCode = Required(variantCode, nameof(variantCode));
            if (!Enum.IsDefined(typeof(EstimateWorkEnvironment), environment))
                throw new ArgumentOutOfRangeException(nameof(environment));
            if (!Enum.IsDefined(typeof(MachineRateAudience), laborAudience))
                throw new ArgumentOutOfRangeException(nameof(laborAudience));
            if (priceProfile.LaborAudience != laborAudience)
                throw new ArgumentException("PriceProfile khong khop doi tuong luong.", nameof(priceProfile));
            Environment = environment;
            LaborAudience = laborAudience;
            PriceProfileId = priceProfile.ProfileId;
            PriceProfileVersion = priceProfile.DataVersion;
            PriceProfileChecksum = priceProfile.Checksum;
            Conditions = new ReadOnlyCollection<string>(
                (conditions ?? Enumerable.Empty<string>())
                    .Select(value => (value ?? string.Empty).Trim())
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToList());
            Bindings = new ReadOnlyCollection<UnitRateResourceBinding>(
                (bindings ?? Enumerable.Empty<UnitRateResourceBinding>())
                    .Where(value => value != null)
                    .OrderBy(value => value.ResourceCode, StringComparer.Ordinal)
                    .ThenBy(value => value.PriceCode, StringComparer.Ordinal)
                    .ToList());
            CanonicalKey = string.Join("|", new[]
            {
                PackageId,
                PackageVersion,
                PackageChecksum,
                NormKey,
                VariantCode,
                ((int)Environment).ToString(CultureInfo.InvariantCulture),
                ((int)LaborAudience).ToString(CultureInfo.InvariantCulture),
                PriceProfileId,
                PriceProfileVersion,
                PriceProfileChecksum,
                string.Join(",", Conditions),
                string.Join(",", Bindings.Select(binding =>
                    binding.ResourceCode + ">" + binding.PriceCode + ">" + binding.Reason))
            });
            RateId = "DG-" + Sha256(CanonicalKey).Substring(0, 16);
        }

        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public string NormKey { get; }
        public string VariantCode { get; }
        public EstimateWorkEnvironment Environment { get; }
        public MachineRateAudience LaborAudience { get; }
        public string PriceProfileId { get; }
        public string PriceProfileVersion { get; }
        public string PriceProfileChecksum { get; }
        public IReadOnlyList<string> Conditions { get; }
        public IReadOnlyList<UnitRateResourceBinding> Bindings { get; }
        public string CanonicalKey { get; }
        public string RateId { get; }

        public bool Equals(EstimateRateIdentity other)
        {
            return other != null && string.Equals(CanonicalKey, other.CanonicalKey, StringComparison.Ordinal);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as EstimateRateIdentity);
        }

        public override int GetHashCode()
        {
            return StringComparer.Ordinal.GetHashCode(CanonicalKey);
        }

        private static string Required(string value, string field)
        {
            string parsed = (value ?? string.Empty).Trim();
            if (parsed.Length == 0)
                throw new ArgumentException(field + " khong duoc trong.", field);
            return parsed;
        }

        private static string Sha256(string value)
        {
            using (SHA256 algorithm = SHA256.Create())
            {
                return string.Concat(algorithm.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))
                    .Select(item => item.ToString("X2", CultureInfo.InvariantCulture)));
            }
        }
    }

    public sealed class EstimateRateGroup
    {
        internal EstimateRateGroup(EstimateRateIdentity identity, IEnumerable<EstimateWorkspaceRow> rows)
        {
            Identity = identity;
            Rows = new ReadOnlyCollection<EstimateWorkspaceRow>(
                rows.OrderBy(row => row.SourceRowHint).ToList());
        }

        public EstimateRateIdentity Identity { get; }
        public IReadOnlyList<EstimateWorkspaceRow> Rows { get; }
    }

    public sealed class EstimateRatePlan
    {
        private EstimateRatePlan(
            IEnumerable<EstimateRateGroup> groups,
            IEnumerable<EstimateWorkspaceRow> textRows,
            IEnumerable<string> errors)
        {
            Groups = new ReadOnlyCollection<EstimateRateGroup>(groups.ToList());
            TextRows = new ReadOnlyCollection<EstimateWorkspaceRow>(textRows.ToList());
            Errors = new ReadOnlyCollection<string>(errors.ToList());
        }

        public IReadOnlyList<EstimateRateGroup> Groups { get; }
        public IReadOnlyList<EstimateWorkspaceRow> TextRows { get; }
        public IReadOnlyList<string> Errors { get; }
        public bool IsValid => Errors.Count == 0;

        public static EstimateRatePlan Build(
            EstimateWorkspace workspace,
            string packageId,
            string packageVersion,
            string packageChecksum,
            IEnumerable<PriceProfile> priceProfiles)
        {
            EstimateWorkspaceValidator.ValidateRequired(workspace);
            var profiles = (priceProfiles ?? Enumerable.Empty<PriceProfile>())
                .Where(profile => profile != null)
                .GroupBy(profile => profile.LaborAudience)
                .ToDictionary(group => group.Key, group => group.ToList());
            var errors = new List<string>();
            foreach (IGrouping<MachineRateAudience, PriceProfile> duplicate in profiles
                .Where(item => item.Value.Count != 1)
                .SelectMany(item => item.Value)
                .GroupBy(item => item.LaborAudience))
            {
                errors.Add("Co nhieu ho so gia cho doi tuong luong " + duplicate.Key + ".");
            }

            var entries = new List<Tuple<EstimateRateIdentity, EstimateWorkspaceRow>>();
            foreach (EstimateWorkspaceRow row in workspace.Rows.Where(item => !item.IsTextRow))
            {
                if (!row.HasCompleteNorm)
                {
                    errors.Add(row.RowId + ": chua chon du ma dinh muc chi tiet.");
                    continue;
                }
                List<PriceProfile> matches;
                if (!profiles.TryGetValue(row.LaborAudience, out matches) || matches.Count != 1)
                {
                    errors.Add(row.RowId + ": chua co dung mot ho so gia cho " + row.LaborAudience + ".");
                    continue;
                }
                entries.Add(Tuple.Create(
                    new EstimateRateIdentity(
                        packageId,
                        packageVersion,
                        packageChecksum,
                        row.NormKey,
                        row.VariantCode,
                        row.Environment,
                        row.LaborAudience,
                        matches[0],
                        row.Conditions,
                        row.Bindings),
                    row));
            }

            EstimateRateGroup[] groups = entries
                .GroupBy(item => item.Item1)
                .Select(group => new EstimateRateGroup(group.Key, group.Select(item => item.Item2)))
                .OrderBy(group => group.Identity.LaborAudience)
                .ThenBy(group => group.Identity.Environment)
                .ThenBy(group => group.Identity.NormKey, StringComparer.Ordinal)
                .ThenBy(group => group.Identity.VariantCode, StringComparer.Ordinal)
                .ToArray();
            return new EstimateRatePlan(
                groups,
                workspace.Rows.Where(row => row.IsTextRow),
                errors.Distinct().ToArray());
        }
    }

    public static class EstimateWorkspaceSerializer
    {
        private const string Magic = "TTBMVN_ESTIMATE_WORKSPACE";

        public static string Serialize(EstimateWorkspace workspace)
        {
            EstimateWorkspaceValidator.ValidateRequired(workspace);
            var lines = new List<string>
            {
                Magic,
                "schema=" + workspace.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                "updatedAt=" + workspace.UpdatedAtUtc.ToString("O", CultureInfo.InvariantCulture),
                "sheetCodeName=" + Encode(workspace.Source.WorksheetCodeName),
                "sheetName=" + Encode(workspace.Source.WorksheetName),
                "address=" + Encode(workspace.Source.SourceAddress),
                "firstRow=" + workspace.Source.FirstDataRow.ToString(CultureInfo.InvariantCulture),
                "lastRow=" + workspace.Source.LastDataRow.ToString(CultureInfo.InvariantCulture),
                "columns=" + string.Join(",", workspace.Source.Columns.AllColumns.Select(
                    value => value.ToString(CultureInfo.InvariantCulture))),
                "rowCount=" + workspace.Rows.Count.ToString(CultureInfo.InvariantCulture)
            };
            for (int index = 0; index < workspace.Rows.Count; index++)
            {
                EstimateWorkspaceRow row = workspace.Rows[index];
                string prefix = "row." + index.ToString(CultureInfo.InvariantCulture) + ".";
                lines.Add(prefix + "id=" + Encode(row.RowId));
                lines.Add(prefix + "binding=" + Encode(row.BindingName));
                lines.Add(prefix + "sourceRow=" + row.SourceRowHint.ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "code=" + Encode(row.WorkCode));
                lines.Add(prefix + "description=" + Encode(row.Description));
                lines.Add(prefix + "unit=" + Encode(row.Unit));
                lines.Add(prefix + "quantityFormula=" + Encode(row.QuantityFormulaLocal));
                lines.Add(prefix + "quantity=" + row.Quantity.ToString("G29", CultureInfo.InvariantCulture));
                lines.Add(prefix + "acceptedFormula=" + Encode(row.AcceptedQuantityFormulaLocal));
                lines.Add(prefix + "accepted=" + row.AcceptedQuantity.ToString("G29", CultureInfo.InvariantCulture));
                lines.Add(prefix + "environment=" + ((int)row.Environment).ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "audience=" + ((int)row.LaborAudience).ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "norm=" + Encode(row.NormKey));
                lines.Add(prefix + "variant=" + Encode(row.VariantCode));
                lines.Add(prefix + "conditions=" + Encode(string.Join("\n", row.Conditions)));
                lines.Add(prefix + "bindingCount=" + row.Bindings.Count.ToString(CultureInfo.InvariantCulture));
                for (int bindingIndex = 0; bindingIndex < row.Bindings.Count; bindingIndex++)
                {
                    UnitRateResourceBinding binding = row.Bindings[bindingIndex];
                    string bindingPrefix = prefix + "binding." +
                        bindingIndex.ToString(CultureInfo.InvariantCulture) + ".";
                    lines.Add(bindingPrefix + "resource=" + Encode(binding.ResourceCode));
                    lines.Add(bindingPrefix + "price=" + Encode(binding.PriceCode));
                    lines.Add(bindingPrefix + "reason=" + Encode(binding.Reason));
                }
            }
            string body = string.Join("\n", lines);
            return body + "\nchecksum=" + Hash(body);
        }

        public static EstimateWorkspace Deserialize(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
                throw new InvalidDataException("EstimateWorkspace payload trong.");
            string[] lines = payload.Replace("\r\n", "\n").Split('\n');
            if (lines.Length < 11 || !string.Equals(lines[0], Magic, StringComparison.Ordinal))
                throw new InvalidDataException("EstimateWorkspace magic khong hop le.");
            string checksumLine = lines[lines.Length - 1];
            if (!checksumLine.StartsWith("checksum=", StringComparison.Ordinal))
                throw new InvalidDataException("EstimateWorkspace thieu checksum.");
            string body = string.Join("\n", lines.Take(lines.Length - 1));
            if (!string.Equals(checksumLine.Substring(9), Hash(body), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("EstimateWorkspace checksum khong khop.");

            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string line in lines.Skip(1).Take(lines.Length - 2))
            {
                int separator = line.IndexOf('=');
                if (separator <= 0 || values.ContainsKey(line.Substring(0, separator)))
                    throw new InvalidDataException("EstimateWorkspace field khong hop le.");
                values.Add(line.Substring(0, separator), line.Substring(separator + 1));
            }
            int schema = Integer(values, "schema");
            if (schema != EstimateWorkspace.CurrentSchemaVersion)
                throw new InvalidDataException("EstimateWorkspace schema khong duoc ho tro: " + schema + ".");
            string[] columnTexts = Required(values, "columns").Split(',');
            if (columnTexts.Length != 11)
                throw new InvalidDataException("EstimateWorkspace columns khong hop le.");
            int[] columns = columnTexts.Select(value => ParseInteger(value, "columns")).ToArray();
            int rowCount = Integer(values, "rowCount");
            if (rowCount < 0 || rowCount > 100000)
                throw new InvalidDataException("EstimateWorkspace rowCount khong hop le.");
            var rows = new List<EstimateWorkspaceRow>(rowCount);
            for (int index = 0; index < rowCount; index++)
            {
                string prefix = "row." + index.ToString(CultureInfo.InvariantCulture) + ".";
                string conditionText = Decode(Required(values, prefix + "conditions"));
                int bindingCount = Integer(values, prefix + "bindingCount");
                if (bindingCount < 0 || bindingCount > 1000)
                    throw new InvalidDataException("EstimateWorkspace bindingCount khong hop le.");
                var bindings = new List<UnitRateResourceBinding>(bindingCount);
                for (int bindingIndex = 0; bindingIndex < bindingCount; bindingIndex++)
                {
                    string bindingPrefix = prefix + "binding." +
                        bindingIndex.ToString(CultureInfo.InvariantCulture) + ".";
                    bindings.Add(new UnitRateResourceBinding(
                        Decode(Required(values, bindingPrefix + "resource")),
                        Decode(Required(values, bindingPrefix + "price")),
                        Decode(Required(values, bindingPrefix + "reason"))));
                }
                rows.Add(new EstimateWorkspaceRow(
                    Decode(Required(values, prefix + "id")),
                    Decode(Required(values, prefix + "binding")),
                    Integer(values, prefix + "sourceRow"),
                    Decode(Required(values, prefix + "code")),
                    Decode(Required(values, prefix + "description")),
                    Decode(Required(values, prefix + "unit")),
                    Decode(Required(values, prefix + "quantityFormula")),
                    Decimal(values, prefix + "quantity"),
                    Decode(Required(values, prefix + "acceptedFormula")),
                    Decimal(values, prefix + "accepted"),
                    (EstimateWorkEnvironment)Integer(values, prefix + "environment"),
                    (MachineRateAudience)Integer(values, prefix + "audience"),
                    Decode(Required(values, prefix + "norm")),
                    Decode(Required(values, prefix + "variant")),
                    conditionText.Length == 0 ? Array.Empty<string>() : conditionText.Split('\n'),
                    bindings));
            }
            DateTime updatedAt;
            if (!DateTime.TryParseExact(
                Required(values, "updatedAt"),
                "O",
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out updatedAt) || updatedAt.Kind != DateTimeKind.Utc)
            {
                throw new InvalidDataException("EstimateWorkspace updatedAt khong hop le.");
            }
            return new EstimateWorkspace(
                new EstimateSourceBinding(
                    Decode(Required(values, "sheetCodeName")),
                    Decode(Required(values, "sheetName")),
                    Decode(Required(values, "address")),
                    Integer(values, "firstRow"),
                    Integer(values, "lastRow"),
                    new EstimateColumnMap(
                        columns[0], columns[1], columns[2], columns[3], columns[4],
                        columns[5], columns[6], columns[7], columns[8], columns[9], columns[10])),
                rows,
                updatedAt);
        }

        private static string Required(IDictionary<string, string> values, string key)
        {
            string value;
            if (!values.TryGetValue(key, out value))
                throw new InvalidDataException("EstimateWorkspace thieu field " + key + ".");
            return value;
        }

        private static int Integer(IDictionary<string, string> values, string key)
        {
            return ParseInteger(Required(values, key), key);
        }

        private static int ParseInteger(string value, string key)
        {
            int parsed;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
                throw new InvalidDataException("EstimateWorkspace " + key + " khong hop le.");
            return parsed;
        }

        private static decimal Decimal(IDictionary<string, string> values, string key)
        {
            decimal parsed;
            string value = Required(values, key);
            if (value.Contains(",") ||
                !decimal.TryParse(value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture, out parsed))
            {
                throw new InvalidDataException("EstimateWorkspace " + key + " khong hop le.");
            }
            return parsed;
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
                throw new InvalidDataException("EstimateWorkspace base64 khong hop le.", ex);
            }
        }

        private static string Hash(string value)
        {
            using (SHA256 algorithm = SHA256.Create())
            {
                return string.Concat(algorithm.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))
                    .Select(item => item.ToString("X2", CultureInfo.InvariantCulture)));
            }
        }
    }
}
