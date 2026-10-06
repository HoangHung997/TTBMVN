using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAddIn1.Core
{
    public enum PriceResourceKind
    {
        Material = 1,
        Labor = 2,
        FuelEnergy = 3,
        MachineOriginalPrice = 4,
        MachineShift = 5
    }

    public enum PriceSourceKind
    {
        Manual = 1,
        WorkbookImport = 2,
        MarketQuote = 3,
        PublishedNotice = 4,
        RegulationReference = 5,
        Calculated = 6
    }

    public sealed class PriceProfileEntry
    {
        public PriceProfileEntry(
            string code,
            PriceResourceKind kind,
            string displayName,
            string unit,
            decimal baseUnitPriceVnd,
            DateTime sourceDate,
            PriceSourceKind sourceKind,
            string sourceReference,
            IEnumerable<string> aliases = null,
            string legacyLookupName = "")
        {
            Code = NormalizeCode(code);
            Kind = kind;
            DisplayName = (displayName ?? string.Empty).Trim();
            Unit = (unit ?? string.Empty).Trim();
            BaseUnitPriceVnd = baseUnitPriceVnd;
            SourceDate = sourceDate.Date;
            SourceKind = sourceKind;
            SourceReference = (sourceReference ?? string.Empty).Trim();
            Aliases = new ReadOnlyCollection<string>(
                (aliases ?? Enumerable.Empty<string>())
                    .Select(NormalizeCode)
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
            LegacyLookupName = (legacyLookupName ?? string.Empty).Trim();
        }

        public string Code { get; }
        public PriceResourceKind Kind { get; }
        public string DisplayName { get; }
        public string Unit { get; }
        public decimal BaseUnitPriceVnd { get; }
        public DateTime SourceDate { get; }
        public PriceSourceKind SourceKind { get; }
        public string SourceReference { get; }
        public IReadOnlyList<string> Aliases { get; }
        public string LegacyLookupName { get; }

        internal static string NormalizeCode(string value)
        {
            return (value ?? string.Empty).Trim();
        }
    }

    public sealed class PriceProfileOverride
    {
        public PriceProfileOverride(
            string code,
            decimal oldUnitPriceVnd,
            decimal newUnitPriceVnd,
            string reason,
            string sourceReference,
            string modifiedBy,
            DateTime modifiedAtUtc)
        {
            Code = PriceProfileEntry.NormalizeCode(code);
            OldUnitPriceVnd = oldUnitPriceVnd;
            NewUnitPriceVnd = newUnitPriceVnd;
            Reason = (reason ?? string.Empty).Trim();
            SourceReference = (sourceReference ?? string.Empty).Trim();
            ModifiedBy = (modifiedBy ?? string.Empty).Trim();
            ModifiedAtUtc = modifiedAtUtc;
        }

        public string Code { get; }
        public decimal OldUnitPriceVnd { get; }
        public decimal NewUnitPriceVnd { get; }
        public string Reason { get; }
        public string SourceReference { get; }
        public string ModifiedBy { get; }
        public DateTime ModifiedAtUtc { get; }
    }

    public sealed class PriceProfilePrice
    {
        internal PriceProfilePrice(PriceProfileEntry entry, PriceProfileOverride priceOverride)
        {
            Entry = entry;
            Override = priceOverride;
        }

        public PriceProfileEntry Entry { get; }
        public PriceProfileOverride Override { get; }
        public decimal AppliedUnitPriceVnd => Override?.NewUnitPriceVnd ?? Entry.BaseUnitPriceVnd;
        public bool IsOverridden => Override != null;
    }

    public sealed class PriceProfile
    {
        public const int CurrentSchemaVersion = 1;
        private readonly IReadOnlyDictionary<string, PriceProfileEntry> byCodeOrAlias;
        private readonly IReadOnlyDictionary<string, PriceProfileOverride> overridesByCode;

        private PriceProfile(
            int schemaVersion,
            string profileId,
            string dataVersion,
            string displayName,
            string location,
            DateTime valuationDate,
            MachineRateAudience laborAudience,
            DateTime createdAtUtc,
            IEnumerable<PriceProfileEntry> entries,
            IEnumerable<PriceProfileOverride> overrides,
            string checksum)
        {
            SchemaVersion = schemaVersion;
            ProfileId = profileId;
            DataVersion = dataVersion;
            DisplayName = displayName;
            Location = location;
            ValuationDate = valuationDate;
            LaborAudience = laborAudience;
            CreatedAtUtc = createdAtUtc;
            Entries = new ReadOnlyCollection<PriceProfileEntry>(
                entries.OrderBy(item => item.Code, StringComparer.OrdinalIgnoreCase).ToList());
            Overrides = new ReadOnlyCollection<PriceProfileOverride>(
                overrides.OrderBy(item => item.Code, StringComparer.OrdinalIgnoreCase).ToList());
            Checksum = checksum;

            var lookup = new Dictionary<string, PriceProfileEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (PriceProfileEntry entry in Entries)
            {
                lookup.Add(entry.Code, entry);
                foreach (string alias in entry.Aliases)
                    lookup.Add(alias, entry);
            }
            byCodeOrAlias = new ReadOnlyDictionary<string, PriceProfileEntry>(lookup);
            overridesByCode = new ReadOnlyDictionary<string, PriceProfileOverride>(
                Overrides.ToDictionary(item => item.Code, StringComparer.OrdinalIgnoreCase));
        }

        public int SchemaVersion { get; }
        public string ProfileId { get; }
        public string DataVersion { get; }
        public string DisplayName { get; }
        public string Location { get; }
        public DateTime ValuationDate { get; }
        public MachineRateAudience LaborAudience { get; }
        public DateTime CreatedAtUtc { get; }
        public string CurrencyCode => "VND";
        public IReadOnlyList<PriceProfileEntry> Entries { get; }
        public IReadOnlyList<PriceProfileOverride> Overrides { get; }
        public string Checksum { get; }

        public static PriceProfile Create(
            string profileId,
            string dataVersion,
            string displayName,
            string location,
            DateTime valuationDate,
            MachineRateAudience laborAudience,
            DateTime createdAtUtc,
            IEnumerable<PriceProfileEntry> entries,
            IEnumerable<PriceProfileOverride> overrides = null)
        {
            var unsigned = new PriceProfile(
                CurrentSchemaVersion,
                (profileId ?? string.Empty).Trim(),
                (dataVersion ?? string.Empty).Trim(),
                (displayName ?? string.Empty).Trim(),
                (location ?? string.Empty).Trim(),
                valuationDate.Date,
                laborAudience,
                NormalizeUtc(createdAtUtc),
                (entries ?? Enumerable.Empty<PriceProfileEntry>()).Where(item => item != null),
                (overrides ?? Enumerable.Empty<PriceProfileOverride>()).Where(item => item != null),
                string.Empty);
            PriceProfileValidator.ValidateRequired(unsigned, verifyChecksum: false);
            string checksum = PriceProfileSerializer.ComputeChecksum(unsigned);
            return Rehydrate(
                unsigned.SchemaVersion,
                unsigned.ProfileId,
                unsigned.DataVersion,
                unsigned.DisplayName,
                unsigned.Location,
                unsigned.ValuationDate,
                unsigned.LaborAudience,
                unsigned.CreatedAtUtc,
                unsigned.Entries,
                unsigned.Overrides,
                checksum);
        }

        internal static PriceProfile Rehydrate(
            int schemaVersion,
            string profileId,
            string dataVersion,
            string displayName,
            string location,
            DateTime valuationDate,
            MachineRateAudience laborAudience,
            DateTime createdAtUtc,
            IEnumerable<PriceProfileEntry> entries,
            IEnumerable<PriceProfileOverride> overrides,
            string checksum)
        {
            return new PriceProfile(
                schemaVersion,
                profileId,
                dataVersion,
                displayName,
                location,
                valuationDate.Date,
                laborAudience,
                NormalizeUtc(createdAtUtc),
                entries ?? Enumerable.Empty<PriceProfileEntry>(),
                overrides ?? Enumerable.Empty<PriceProfileOverride>(),
                (checksum ?? string.Empty).Trim().ToUpperInvariant());
        }

        public bool TryFind(string codeOrAlias, out PriceProfilePrice price)
        {
            PriceProfileEntry entry;
            if (!byCodeOrAlias.TryGetValue(PriceProfileEntry.NormalizeCode(codeOrAlias), out entry))
            {
                price = null;
                return false;
            }
            PriceProfileOverride value;
            overridesByCode.TryGetValue(entry.Code, out value);
            price = new PriceProfilePrice(entry, value);
            return true;
        }

        public PriceProfilePrice FindRequired(string codeOrAlias)
        {
            PriceProfilePrice value;
            if (!TryFind(codeOrAlias, out value))
                throw new KeyNotFoundException("Thieu gia cho ma: " + codeOrAlias + ".");
            return value;
        }

        public MachineRatePriceProfile ToMachineRatePriceProfile()
        {
            return new MachineRatePriceProfile(
                FlattenPrices(PriceResourceKind.FuelEnergy),
                FlattenPrices(PriceResourceKind.Labor));
        }

        private IEnumerable<KeyValuePair<string, decimal>> FlattenPrices(PriceResourceKind kind)
        {
            foreach (PriceProfileEntry entry in Entries.Where(item => item.Kind == kind))
            {
                PriceProfilePrice value = FindRequired(entry.Code);
                yield return new KeyValuePair<string, decimal>(entry.Code, value.AppliedUnitPriceVnd);
                foreach (string alias in entry.Aliases)
                    yield return new KeyValuePair<string, decimal>(alias, value.AppliedUnitPriceVnd);
            }
        }

        private static DateTime NormalizeUtc(DateTime value)
        {
            if (value.Kind == DateTimeKind.Utc)
                return value;
            if (value.Kind == DateTimeKind.Local)
                return value.ToUniversalTime();
            return DateTime.SpecifyKind(value, DateTimeKind.Utc);
        }
    }

    public sealed class PriceProfileValidationResult
    {
        internal PriceProfileValidationResult(IEnumerable<string> errors)
        {
            Errors = new ReadOnlyCollection<string>(errors.ToList());
        }

        public IReadOnlyList<string> Errors { get; }
        public bool IsValid => Errors.Count == 0;
    }

    public static class PriceProfileValidator
    {
        private static readonly Regex IdPattern = new Regex(
            @"^[A-Za-z0-9][A-Za-z0-9._-]{0,159}$",
            RegexOptions.CultureInvariant);
        private static readonly Regex VersionPattern = new Regex(
            @"^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$",
            RegexOptions.CultureInvariant);
        private static readonly Regex Sha256Pattern = new Regex(
            "^[A-F0-9]{64}$",
            RegexOptions.CultureInvariant);

        public static PriceProfileValidationResult Validate(PriceProfile profile)
        {
            return ValidateInternal(profile, verifyChecksum: true);
        }

        internal static void ValidateRequired(PriceProfile profile, bool verifyChecksum)
        {
            PriceProfileValidationResult result = ValidateInternal(profile, verifyChecksum);
            if (!result.IsValid)
                throw new ArgumentException(string.Join(" ", result.Errors), nameof(profile));
        }

        private static PriceProfileValidationResult ValidateInternal(
            PriceProfile profile,
            bool verifyChecksum)
        {
            var errors = new List<string>();
            if (profile == null)
            {
                errors.Add("PriceProfile khong duoc null.");
                return new PriceProfileValidationResult(errors);
            }
            if (profile.SchemaVersion != PriceProfile.CurrentSchemaVersion)
                errors.Add("PriceProfile schema khong duoc ho tro.");
            if (!IdPattern.IsMatch(profile.ProfileId ?? string.Empty))
                errors.Add("PriceProfileId khong hop le.");
            if (!VersionPattern.IsMatch(profile.DataVersion ?? string.Empty))
                errors.Add("PriceProfile DataVersion khong hop le.");
            RequiredText(profile.DisplayName, "DisplayName", 240, errors);
            RequiredText(profile.Location, "Location", 240, errors);
            if (profile.ValuationDate.Year < 2000 || profile.ValuationDate.TimeOfDay != TimeSpan.Zero)
                errors.Add("ValuationDate khong hop le.");
            if (!Enum.IsDefined(typeof(MachineRateAudience), profile.LaborAudience))
                errors.Add("LaborAudience khong hop le.");
            if (profile.CreatedAtUtc.Kind != DateTimeKind.Utc)
                errors.Add("CreatedAtUtc phai la UTC.");
            if (profile.Entries == null || profile.Entries.Count == 0 || profile.Entries.Count > 10000)
                errors.Add("PriceProfile phai co tu 1 den 10000 entry.");

            var lookupKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var entries = new Dictionary<string, PriceProfileEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (PriceProfileEntry entry in profile.Entries ?? Array.Empty<PriceProfileEntry>())
            {
                if (entry == null)
                {
                    errors.Add("PriceProfile co entry null.");
                    continue;
                }
                RequiredText(entry.Code, "Entry.Code", 160, errors);
                RequiredText(entry.DisplayName, "Entry.DisplayName", 300, errors);
                RequiredText(entry.Unit, "Entry.Unit", 80, errors);
                RequiredText(entry.SourceReference, "Entry.SourceReference", 1000, errors);
                if (!Enum.IsDefined(typeof(PriceResourceKind), entry.Kind))
                    errors.Add(entry.Code + ": Kind khong hop le.");
                if (!Enum.IsDefined(typeof(PriceSourceKind), entry.SourceKind))
                    errors.Add(entry.Code + ": SourceKind khong hop le.");
                if (entry.BaseUnitPriceVnd < 0m)
                    errors.Add(entry.Code + ": gia goc khong duoc am.");
                if (entry.SourceDate.Year < 2000 || entry.SourceDate > profile.ValuationDate)
                    errors.Add(entry.Code + ": SourceDate vuot ngay gia hoac khong hop le.");
                if ((entry.LegacyLookupName ?? string.Empty).Length > 300)
                    errors.Add(entry.Code + ": LegacyLookupName qua dai.");
                if (!lookupKeys.Add(entry.Code))
                    errors.Add("Ma/alias gia bi trung: " + entry.Code + ".");
                else
                    entries[entry.Code] = entry;
                foreach (string alias in entry.Aliases ?? Array.Empty<string>())
                {
                    RequiredText(alias, entry.Code + ".Alias", 160, errors);
                    if (!lookupKeys.Add(alias))
                        errors.Add("Ma/alias gia bi trung: " + alias + ".");
                }
            }

            var overrideCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PriceProfileOverride value in profile.Overrides ?? Array.Empty<PriceProfileOverride>())
            {
                if (value == null)
                {
                    errors.Add("PriceProfile co override null.");
                    continue;
                }
                PriceProfileEntry entry;
                if (!entries.TryGetValue(value.Code, out entry))
                    errors.Add("Override khong tro dung ma entry: " + value.Code + ".");
                else if (value.OldUnitPriceVnd != entry.BaseUnitPriceVnd)
                    errors.Add(value.Code + ": gia cu override khong khop gia goc.");
                if (!overrideCodes.Add(value.Code))
                    errors.Add("Override bi trung: " + value.Code + ".");
                if (value.OldUnitPriceVnd < 0m || value.NewUnitPriceVnd < 0m)
                    errors.Add(value.Code + ": gia override khong duoc am.");
                RequiredText(value.Reason, value.Code + ".Reason", 1024, errors);
                RequiredText(value.SourceReference, value.Code + ".OverrideSource", 1000, errors);
                RequiredText(value.ModifiedBy, value.Code + ".ModifiedBy", 200, errors);
                if (value.ModifiedAtUtc.Kind != DateTimeKind.Utc ||
                    value.ModifiedAtUtc > profile.CreatedAtUtc)
                {
                    errors.Add(value.Code + ": ModifiedAtUtc khong hop le.");
                }
            }

            if (verifyChecksum)
            {
                if (!Sha256Pattern.IsMatch(profile.Checksum ?? string.Empty))
                    errors.Add("PriceProfile checksum khong hop le.");
                else if (errors.Count == 0 &&
                    !string.Equals(
                        PriceProfileSerializer.ComputeChecksum(profile),
                        profile.Checksum,
                        StringComparison.Ordinal))
                {
                    errors.Add("PriceProfile checksum khong khop noi dung.");
                }
            }
            return new PriceProfileValidationResult(errors);
        }

        private static void RequiredText(
            string value,
            string field,
            int maximumLength,
            IList<string> errors)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length > maximumLength)
                errors.Add(field + " phai co tu 1 den " + maximumLength + " ky tu.");
        }
    }

    public sealed class PriceRequirement
    {
        public PriceRequirement(string code, PriceResourceKind kind, string unit)
        {
            Code = PriceProfileEntry.NormalizeCode(code);
            Kind = kind;
            Unit = (unit ?? string.Empty).Trim();
        }

        public string Code { get; }
        public PriceResourceKind Kind { get; }
        public string Unit { get; }
    }

    public enum PriceCoverageIssueCode
    {
        Missing = 1,
        KindMismatch = 2,
        UnitMismatch = 3,
        InvalidRequirement = 4
    }

    public sealed class PriceCoverageIssue
    {
        internal PriceCoverageIssue(PriceCoverageIssueCode issueCode, string code, string message)
        {
            IssueCode = issueCode;
            Code = code;
            Message = message;
        }

        public PriceCoverageIssueCode IssueCode { get; }
        public string Code { get; }
        public string Message { get; }
    }

    public sealed class PriceCoverageResult
    {
        internal PriceCoverageResult(IEnumerable<PriceCoverageIssue> issues)
        {
            Issues = new ReadOnlyCollection<PriceCoverageIssue>(issues.ToList());
        }

        public IReadOnlyList<PriceCoverageIssue> Issues { get; }
        public bool IsComplete => Issues.Count == 0;
    }

    public static class PriceProfileCoverageValidator
    {
        public static PriceCoverageResult Validate(
            PriceProfile profile,
            IEnumerable<PriceRequirement> requirements)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            var issues = new List<PriceCoverageIssue>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PriceRequirement requirement in requirements ?? Enumerable.Empty<PriceRequirement>())
            {
                if (requirement == null ||
                    string.IsNullOrWhiteSpace(requirement.Code) ||
                    string.IsNullOrWhiteSpace(requirement.Unit) ||
                    !Enum.IsDefined(typeof(PriceResourceKind), requirement.Kind) ||
                    !seen.Add(requirement.Code))
                {
                    issues.Add(new PriceCoverageIssue(
                        PriceCoverageIssueCode.InvalidRequirement,
                        requirement?.Code ?? string.Empty,
                        "Yeu cau gia khong hop le hoac bi trung."));
                    continue;
                }
                PriceProfilePrice price;
                if (!profile.TryFind(requirement.Code, out price))
                {
                    issues.Add(new PriceCoverageIssue(
                        PriceCoverageIssueCode.Missing,
                        requirement.Code,
                        "Thieu gia " + requirement.Code + "."));
                    continue;
                }
                if (price.Entry.Kind != requirement.Kind)
                {
                    issues.Add(new PriceCoverageIssue(
                        PriceCoverageIssueCode.KindMismatch,
                        requirement.Code,
                        "Sai loai gia " + requirement.Code + "."));
                }
                if (!string.Equals(price.Entry.Unit, requirement.Unit, StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add(new PriceCoverageIssue(
                        PriceCoverageIssueCode.UnitMismatch,
                        requirement.Code,
                        "Sai don vi gia " + requirement.Code + "."));
                }
            }
            return new PriceCoverageResult(issues);
        }
    }

    public static class PriceProfileSerializer
    {
        private const string Magic = "TTBMVN_PRICE_PROFILE";

        public static string Serialize(PriceProfile profile)
        {
            PriceProfileValidator.ValidateRequired(profile, verifyChecksum: true);
            return BuildBody(profile) + "\nchecksum=" + profile.Checksum;
        }

        public static PriceProfile Deserialize(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
                throw new InvalidDataException("PriceProfile payload trong.");
            string[] lines = payload.Replace("\r\n", "\n").Split('\n');
            if (lines.Length < 3 || !string.Equals(lines[0], Magic, StringComparison.Ordinal))
                throw new InvalidDataException("PriceProfile magic khong hop le.");
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int index = 1; index < lines.Length; index++)
            {
                int separator = lines[index].IndexOf('=');
                if (separator <= 0 || values.ContainsKey(lines[index].Substring(0, separator)))
                    throw new InvalidDataException("PriceProfile field khong hop le.");
                values.Add(
                    lines[index].Substring(0, separator),
                    lines[index].Substring(separator + 1));
            }

            var consumed = new HashSet<string>(StringComparer.Ordinal);
            int schema = ParseInt(Get(values, consumed, "schema"), "schema", 1, int.MaxValue);
            int entryCount = ParseInt(Get(values, consumed, "entryCount"), "entryCount", 1, 10000);
            var entries = new List<PriceProfileEntry>(entryCount);
            for (int index = 0; index < entryCount; index++)
            {
                string prefix = "entry." + index.ToString(CultureInfo.InvariantCulture) + ".";
                int aliasCount = ParseInt(Get(values, consumed, prefix + "aliasCount"), prefix + "aliasCount", 0, 1000);
                var aliases = new List<string>(aliasCount);
                for (int aliasIndex = 0; aliasIndex < aliasCount; aliasIndex++)
                {
                    aliases.Add(Decode(Get(
                        values,
                        consumed,
                        prefix + "alias." + aliasIndex.ToString(CultureInfo.InvariantCulture))));
                }
                entries.Add(new PriceProfileEntry(
                    Decode(Get(values, consumed, prefix + "code")),
                    (PriceResourceKind)ParseInt(Get(values, consumed, prefix + "kind"), prefix + "kind", 1, 5),
                    Decode(Get(values, consumed, prefix + "name")),
                    Decode(Get(values, consumed, prefix + "unit")),
                    ParseDecimal(Get(values, consumed, prefix + "basePrice"), prefix + "basePrice"),
                    ParseDate(Get(values, consumed, prefix + "sourceDate"), prefix + "sourceDate"),
                    (PriceSourceKind)ParseInt(Get(values, consumed, prefix + "sourceKind"), prefix + "sourceKind", 1, 6),
                    Decode(Get(values, consumed, prefix + "sourceReference")),
                    aliases,
                    Decode(Get(values, consumed, prefix + "legacyLookupName"))));
            }
            int overrideCount = ParseInt(Get(values, consumed, "overrideCount"), "overrideCount", 0, 10000);
            var overrides = new List<PriceProfileOverride>(overrideCount);
            for (int index = 0; index < overrideCount; index++)
            {
                string prefix = "override." + index.ToString(CultureInfo.InvariantCulture) + ".";
                overrides.Add(new PriceProfileOverride(
                    Decode(Get(values, consumed, prefix + "code")),
                    ParseDecimal(Get(values, consumed, prefix + "oldPrice"), prefix + "oldPrice"),
                    ParseDecimal(Get(values, consumed, prefix + "newPrice"), prefix + "newPrice"),
                    Decode(Get(values, consumed, prefix + "reason")),
                    Decode(Get(values, consumed, prefix + "sourceReference")),
                    Decode(Get(values, consumed, prefix + "modifiedBy")),
                    ParseUtc(Get(values, consumed, prefix + "modifiedAtUtc"), prefix + "modifiedAtUtc")));
            }
            PriceProfile profile = PriceProfile.Rehydrate(
                schema,
                Decode(Get(values, consumed, "profileId")),
                Decode(Get(values, consumed, "dataVersion")),
                Decode(Get(values, consumed, "displayName")),
                Decode(Get(values, consumed, "location")),
                ParseDate(Get(values, consumed, "valuationDate"), "valuationDate"),
                (MachineRateAudience)ParseInt(Get(values, consumed, "laborAudience"), "laborAudience", 1, 2),
                ParseUtc(Get(values, consumed, "createdAtUtc"), "createdAtUtc"),
                entries,
                overrides,
                Get(values, consumed, "checksum"));
            if (consumed.Count != values.Count)
                throw new InvalidDataException("PriceProfile co field khong duoc ho tro.");
            PriceProfileValidationResult validation = PriceProfileValidator.Validate(profile);
            if (!validation.IsValid)
                throw new InvalidDataException(string.Join(" ", validation.Errors));
            return profile;
        }

        public static string ComputeChecksum(PriceProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            return Sha256(BuildBody(profile));
        }

        public static string ComputePayloadChecksum(string payload)
        {
            return Sha256(payload ?? string.Empty);
        }

        private static string BuildBody(PriceProfile profile)
        {
            var lines = new List<string>
            {
                Magic,
                "schema=" + profile.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                "profileId=" + Encode(profile.ProfileId),
                "dataVersion=" + Encode(profile.DataVersion),
                "displayName=" + Encode(profile.DisplayName),
                "location=" + Encode(profile.Location),
                "valuationDate=" + profile.ValuationDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                "laborAudience=" + ((int)profile.LaborAudience).ToString(CultureInfo.InvariantCulture),
                "createdAtUtc=" + profile.CreatedAtUtc.ToString("O", CultureInfo.InvariantCulture),
                "entryCount=" + profile.Entries.Count.ToString(CultureInfo.InvariantCulture)
            };
            for (int index = 0; index < profile.Entries.Count; index++)
            {
                PriceProfileEntry entry = profile.Entries[index];
                string prefix = "entry." + index.ToString(CultureInfo.InvariantCulture) + ".";
                lines.Add(prefix + "code=" + Encode(entry.Code));
                lines.Add(prefix + "kind=" + ((int)entry.Kind).ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "name=" + Encode(entry.DisplayName));
                lines.Add(prefix + "unit=" + Encode(entry.Unit));
                lines.Add(prefix + "basePrice=" + entry.BaseUnitPriceVnd.ToString("G29", CultureInfo.InvariantCulture));
                lines.Add(prefix + "sourceDate=" + entry.SourceDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                lines.Add(prefix + "sourceKind=" + ((int)entry.SourceKind).ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "sourceReference=" + Encode(entry.SourceReference));
                lines.Add(prefix + "legacyLookupName=" + Encode(entry.LegacyLookupName));
                lines.Add(prefix + "aliasCount=" + entry.Aliases.Count.ToString(CultureInfo.InvariantCulture));
                for (int aliasIndex = 0; aliasIndex < entry.Aliases.Count; aliasIndex++)
                {
                    lines.Add(prefix + "alias." + aliasIndex.ToString(CultureInfo.InvariantCulture) +
                        "=" + Encode(entry.Aliases[aliasIndex]));
                }
            }
            lines.Add("overrideCount=" + profile.Overrides.Count.ToString(CultureInfo.InvariantCulture));
            for (int index = 0; index < profile.Overrides.Count; index++)
            {
                PriceProfileOverride value = profile.Overrides[index];
                string prefix = "override." + index.ToString(CultureInfo.InvariantCulture) + ".";
                lines.Add(prefix + "code=" + Encode(value.Code));
                lines.Add(prefix + "oldPrice=" + value.OldUnitPriceVnd.ToString("G29", CultureInfo.InvariantCulture));
                lines.Add(prefix + "newPrice=" + value.NewUnitPriceVnd.ToString("G29", CultureInfo.InvariantCulture));
                lines.Add(prefix + "reason=" + Encode(value.Reason));
                lines.Add(prefix + "sourceReference=" + Encode(value.SourceReference));
                lines.Add(prefix + "modifiedBy=" + Encode(value.ModifiedBy));
                lines.Add(prefix + "modifiedAtUtc=" + value.ModifiedAtUtc.ToString("O", CultureInfo.InvariantCulture));
            }
            return string.Join("\n", lines);
        }

        private static string Get(
            IDictionary<string, string> values,
            ISet<string> consumed,
            string key)
        {
            string value;
            if (!values.TryGetValue(key, out value))
                throw new InvalidDataException("PriceProfile thieu field " + key + ".");
            consumed.Add(key);
            return value;
        }

        private static int ParseInt(string value, string field, int minimum, int maximum)
        {
            int parsed;
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out parsed) ||
                parsed < minimum || parsed > maximum)
            {
                throw new InvalidDataException("PriceProfile " + field + " khong hop le.");
            }
            return parsed;
        }

        private static decimal ParseDecimal(string value, string field)
        {
            decimal parsed;
            if (value.Contains(",") ||
                !decimal.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out parsed))
            {
                throw new InvalidDataException("PriceProfile " + field + " khong hop le.");
            }
            return parsed;
        }

        private static DateTime ParseDate(string value, string field)
        {
            DateTime parsed;
            if (!DateTime.TryParseExact(
                value,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out parsed))
            {
                throw new InvalidDataException("PriceProfile " + field + " khong hop le.");
            }
            return parsed.Date;
        }

        private static DateTime ParseUtc(string value, string field)
        {
            DateTime parsed;
            if (!DateTime.TryParseExact(
                value,
                "O",
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out parsed) || parsed.Kind != DateTimeKind.Utc)
            {
                throw new InvalidDataException("PriceProfile " + field + " khong hop le.");
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
                throw new InvalidDataException("PriceProfile base64 khong hop le.", ex);
            }
        }

        private static string Sha256(string value)
        {
            using (SHA256 algorithm = SHA256.Create())
            {
                byte[] hash = algorithm.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty));
                return string.Concat(hash.Select(item => item.ToString("X2", CultureInfo.InvariantCulture)));
            }
        }
    }
}
