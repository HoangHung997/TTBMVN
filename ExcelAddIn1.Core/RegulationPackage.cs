using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelAddIn1.Core
{
    public enum RegulationPackageStatus
    {
        Draft = 0,
        Published = 1,
        Superseded = 2,
        Withdrawn = 3
    }

    public enum RegulationModuleKind
    {
        TechnicalProcess = 1,
        Norm = 2,
        CostRule = 3,
        MachineRate = 4,
        Geography = 5,
        Compliance = 6
    }

    public sealed class RegulationPackageSourceDocument
    {
        public RegulationPackageSourceDocument(
            string documentId,
            string title,
            string publisher,
            DateTime issuedDate,
            DateTime effectiveFrom,
            DateTime? effectiveTo,
            string officialUri,
            string contentChecksum)
        {
            DocumentId = documentId;
            Title = title;
            Publisher = publisher;
            IssuedDate = issuedDate;
            EffectiveFrom = effectiveFrom;
            EffectiveTo = effectiveTo;
            OfficialUri = officialUri;
            ContentChecksum = contentChecksum;
        }

        public string DocumentId { get; }
        public string Title { get; }
        public string Publisher { get; }
        public DateTime IssuedDate { get; }
        public DateTime EffectiveFrom { get; }
        public DateTime? EffectiveTo { get; }
        public string OfficialUri { get; }
        public string ContentChecksum { get; }
    }

    public sealed class RegulationPackageModuleManifest
    {
        public RegulationPackageModuleManifest(
            RegulationModuleKind kind,
            string moduleId,
            int schemaVersion,
            string dataVersion,
            int recordCount,
            string contentChecksum)
        {
            Kind = kind;
            ModuleId = moduleId;
            SchemaVersion = schemaVersion;
            DataVersion = dataVersion;
            RecordCount = recordCount;
            ContentChecksum = contentChecksum;
        }

        public RegulationModuleKind Kind { get; }
        public string ModuleId { get; }
        public int SchemaVersion { get; }
        public string DataVersion { get; }
        public int RecordCount { get; }
        public string ContentChecksum { get; }
    }

    public sealed class RegulationPackage
    {
        public const int CurrentSchemaVersion = 1;

        private RegulationPackage(
            int schemaVersion,
            string packageId,
            string dataVersion,
            DateTime effectiveFrom,
            DateTime? effectiveTo,
            RegulationPackageStatus status,
            string transitionNote,
            IEnumerable<RegulationPackageSourceDocument> sources,
            IEnumerable<RegulationPackageModuleManifest> modules,
            string packageChecksum)
        {
            SchemaVersion = schemaVersion;
            PackageId = packageId;
            DataVersion = dataVersion;
            EffectiveFrom = effectiveFrom;
            EffectiveTo = effectiveTo;
            Status = status;
            TransitionNote = transitionNote;
            Sources = new ReadOnlyCollection<RegulationPackageSourceDocument>(
                (sources ?? Enumerable.Empty<RegulationPackageSourceDocument>()).ToList());
            Modules = new ReadOnlyCollection<RegulationPackageModuleManifest>(
                (modules ?? Enumerable.Empty<RegulationPackageModuleManifest>()).ToList());
            PackageChecksum = packageChecksum;
        }

        public int SchemaVersion { get; }
        public string PackageId { get; }
        public string DataVersion { get; }
        public DateTime EffectiveFrom { get; }
        public DateTime? EffectiveTo { get; }
        public RegulationPackageStatus Status { get; }
        public string TransitionNote { get; }
        public IReadOnlyList<RegulationPackageSourceDocument> Sources { get; }
        public IReadOnlyList<RegulationPackageModuleManifest> Modules { get; }
        public string PackageChecksum { get; }

        public static RegulationPackage Create(
            string packageId,
            string dataVersion,
            DateTime effectiveFrom,
            DateTime? effectiveTo,
            RegulationPackageStatus status,
            string transitionNote,
            IEnumerable<RegulationPackageSourceDocument> sources,
            IEnumerable<RegulationPackageModuleManifest> modules)
        {
            var unsigned = new RegulationPackage(
                CurrentSchemaVersion,
                packageId,
                dataVersion,
                effectiveFrom,
                effectiveTo,
                status,
                transitionNote,
                sources,
                modules,
                string.Empty);
            RegulationPackageValidationResult validation =
                RegulationPackageValidator.ValidateForSealing(unsigned);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(packageId));

            RegulationPackage normalized = new RegulationPackage(
                unsigned.SchemaVersion,
                unsigned.PackageId.Trim(),
                unsigned.DataVersion.Trim(),
                unsigned.EffectiveFrom.Date,
                unsigned.EffectiveTo?.Date,
                unsigned.Status,
                unsigned.TransitionNote ?? string.Empty,
                NormalizeSources(unsigned.Sources)
                    .OrderBy(source => source.DocumentId, StringComparer.Ordinal),
                NormalizeModules(unsigned.Modules)
                    .OrderBy(module => (int)module.Kind)
                    .ThenBy(module => module.ModuleId, StringComparer.Ordinal),
                string.Empty);
            return new RegulationPackage(
                normalized.SchemaVersion,
                normalized.PackageId,
                normalized.DataVersion,
                normalized.EffectiveFrom,
                normalized.EffectiveTo,
                normalized.Status,
                normalized.TransitionNote,
                normalized.Sources,
                normalized.Modules,
                RegulationPackageSerializer.ComputePackageChecksum(normalized));
        }

        internal static RegulationPackage Rehydrate(
            int schemaVersion,
            string packageId,
            string dataVersion,
            DateTime effectiveFrom,
            DateTime? effectiveTo,
            RegulationPackageStatus status,
            string transitionNote,
            IEnumerable<RegulationPackageSourceDocument> sources,
            IEnumerable<RegulationPackageModuleManifest> modules,
            string packageChecksum)
        {
            return new RegulationPackage(
                schemaVersion,
                packageId,
                dataVersion,
                effectiveFrom,
                effectiveTo,
                status,
                transitionNote,
                sources,
                modules,
                packageChecksum);
        }

        private static IEnumerable<RegulationPackageSourceDocument> NormalizeSources(
            IEnumerable<RegulationPackageSourceDocument> sources)
        {
            foreach (RegulationPackageSourceDocument source in sources)
            {
                yield return new RegulationPackageSourceDocument(
                    source.DocumentId.Trim(),
                    source.Title.Trim(),
                    source.Publisher.Trim(),
                    source.IssuedDate.Date,
                    source.EffectiveFrom.Date,
                    source.EffectiveTo?.Date,
                    source.OfficialUri.Trim(),
                    source.ContentChecksum.Trim().ToUpperInvariant());
            }
        }

        private static IEnumerable<RegulationPackageModuleManifest> NormalizeModules(
            IEnumerable<RegulationPackageModuleManifest> modules)
        {
            foreach (RegulationPackageModuleManifest module in modules)
            {
                yield return new RegulationPackageModuleManifest(
                    module.Kind,
                    module.ModuleId.Trim(),
                    module.SchemaVersion,
                    module.DataVersion.Trim(),
                    module.RecordCount,
                    module.ContentChecksum.Trim().ToUpperInvariant());
            }
        }
    }

    public sealed class RegulationPackageValidationResult
    {
        internal RegulationPackageValidationResult(IEnumerable<string> errors)
        {
            Errors = new ReadOnlyCollection<string>(
                (errors ?? Enumerable.Empty<string>()).ToList());
        }

        public bool IsValid => Errors.Count == 0;
        public IReadOnlyList<string> Errors { get; }
    }

    public static class RegulationPackageValidator
    {
        private static readonly Regex IdPattern = new Regex(
            "^[A-Za-z0-9][A-Za-z0-9._-]{2,159}$",
            RegexOptions.CultureInvariant);
        private static readonly Regex VersionPattern = new Regex(
            "^[0-9]+\\.[0-9]+\\.[0-9]+(?:-[0-9A-Za-z.-]+)?$",
            RegexOptions.CultureInvariant);
        private static readonly Regex Sha256Pattern = new Regex(
            "^[0-9A-Fa-f]{64}$",
            RegexOptions.CultureInvariant);

        public static RegulationPackageValidationResult Validate(RegulationPackage package)
        {
            return Validate(package, true);
        }

        internal static RegulationPackageValidationResult ValidateForSealing(RegulationPackage package)
        {
            return Validate(package, false);
        }

        private static RegulationPackageValidationResult Validate(
            RegulationPackage package,
            bool verifyPackageChecksum)
        {
            var errors = new List<string>();
            if (package == null)
            {
                errors.Add("Regulation package khong duoc null.");
                return new RegulationPackageValidationResult(errors);
            }

            if (package.SchemaVersion != RegulationPackage.CurrentSchemaVersion)
                errors.Add("Package schema version khong duoc ho tro: " + package.SchemaVersion + ".");
            ValidateId(package.PackageId, "PackageId", errors);
            ValidateVersion(package.DataVersion, "DataVersion", errors);
            ValidateDate(package.EffectiveFrom, "EffectiveFrom", errors);
            if (package.EffectiveTo.HasValue)
            {
                ValidateDate(package.EffectiveTo.Value, "EffectiveTo", errors);
                if (package.EffectiveTo.Value.Date < package.EffectiveFrom.Date)
                    errors.Add("EffectiveTo khong duoc truoc EffectiveFrom.");
            }
            if (!Enum.IsDefined(typeof(RegulationPackageStatus), package.Status))
                errors.Add("Package status khong hop le.");
            if ((package.TransitionNote ?? string.Empty).Length > 4096)
                errors.Add("TransitionNote khong duoc vuot qua 4096 ky tu.");

            ValidateSources(package.Sources, errors);
            ValidateModules(package.Modules, package.Status, errors);

            if (package.Status != RegulationPackageStatus.Draft && package.Sources.Count == 0)
                errors.Add("Package da phat hanh phai co it nhat mot van ban nguon.");

            if (verifyPackageChecksum)
            {
                if (!IsSha256(package.PackageChecksum))
                {
                    errors.Add("PackageChecksum phai la SHA-256 hex 64 ky tu.");
                }
                else if (errors.Count == 0)
                {
                    string actual = RegulationPackageSerializer.ComputePackageChecksum(package);
                    if (!string.Equals(actual, package.PackageChecksum, StringComparison.OrdinalIgnoreCase))
                        errors.Add("PackageChecksum khong khop noi dung manifest.");
                }
            }

            return new RegulationPackageValidationResult(errors);
        }

        private static void ValidateSources(
            IReadOnlyList<RegulationPackageSourceDocument> sources,
            List<string> errors)
        {
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < sources.Count; index++)
            {
                RegulationPackageSourceDocument source = sources[index];
                string prefix = "Sources[" + index + "]";
                if (source == null)
                {
                    errors.Add(prefix + " khong duoc null.");
                    continue;
                }

                ValidateId(source.DocumentId, prefix + ".DocumentId", errors);
                if (!string.IsNullOrWhiteSpace(source.DocumentId) && !ids.Add(source.DocumentId.Trim()))
                    errors.Add(prefix + ".DocumentId bi trung.");
                ValidateRequiredText(source.Title, prefix + ".Title", 512, errors);
                ValidateRequiredText(source.Publisher, prefix + ".Publisher", 256, errors);
                ValidateDate(source.IssuedDate, prefix + ".IssuedDate", errors);
                ValidateDate(source.EffectiveFrom, prefix + ".EffectiveFrom", errors);
                if (source.EffectiveTo.HasValue)
                {
                    ValidateDate(source.EffectiveTo.Value, prefix + ".EffectiveTo", errors);
                    if (source.EffectiveTo.Value.Date < source.EffectiveFrom.Date)
                        errors.Add(prefix + ".EffectiveTo khong duoc truoc EffectiveFrom.");
                }

                Uri uri;
                if (!Uri.TryCreate(source.OfficialUri, UriKind.Absolute, out uri) ||
                    (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp))
                {
                    errors.Add(prefix + ".OfficialUri phai la URL HTTP/HTTPS tuyet doi.");
                }
                if (!IsSha256(source.ContentChecksum))
                    errors.Add(prefix + ".ContentChecksum phai la SHA-256 hex 64 ky tu.");
            }
        }

        private static void ValidateModules(
            IReadOnlyList<RegulationPackageModuleManifest> modules,
            RegulationPackageStatus status,
            List<string> errors)
        {
            var kinds = new HashSet<RegulationModuleKind>();
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < modules.Count; index++)
            {
                RegulationPackageModuleManifest module = modules[index];
                string prefix = "Modules[" + index + "]";
                if (module == null)
                {
                    errors.Add(prefix + " khong duoc null.");
                    continue;
                }

                if (!Enum.IsDefined(typeof(RegulationModuleKind), module.Kind))
                    errors.Add(prefix + ".Kind khong hop le.");
                else if (!kinds.Add(module.Kind))
                    errors.Add(prefix + ".Kind bi trung.");
                ValidateId(module.ModuleId, prefix + ".ModuleId", errors);
                if (!string.IsNullOrWhiteSpace(module.ModuleId) && !ids.Add(module.ModuleId.Trim()))
                    errors.Add(prefix + ".ModuleId bi trung.");
                if (module.SchemaVersion <= 0)
                    errors.Add(prefix + ".SchemaVersion phai lon hon 0.");
                ValidateVersion(module.DataVersion, prefix + ".DataVersion", errors);
                if (module.RecordCount < 0)
                    errors.Add(prefix + ".RecordCount khong duoc am.");
                if (!IsSha256(module.ContentChecksum))
                    errors.Add(prefix + ".ContentChecksum phai la SHA-256 hex 64 ky tu.");
            }

            if (status == RegulationPackageStatus.Draft)
                return;

            foreach (RegulationModuleKind required in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                if (!kinds.Contains(required))
                    errors.Add("Package da phat hanh thieu module " + required + ".");
            }
        }

        private static void ValidateId(string value, string fieldName, List<string> errors)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (!IdPattern.IsMatch(normalized))
                errors.Add(fieldName + " khong dung dinh dang ID.");
        }

        private static void ValidateVersion(string value, string fieldName, List<string> errors)
        {
            if (!IsDataVersion(value))
                errors.Add(fieldName + " phai dung dang major.minor.patch.");
        }

        internal static bool IsDataVersion(string value)
        {
            return VersionPattern.IsMatch((value ?? string.Empty).Trim());
        }

        private static void ValidateRequiredText(
            string value,
            string fieldName,
            int maximumLength,
            List<string> errors)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length == 0)
                errors.Add(fieldName + " la bat buoc.");
            else if (normalized.Length > maximumLength)
                errors.Add(fieldName + " khong duoc vuot qua " + maximumLength + " ky tu.");
        }

        private static void ValidateDate(DateTime value, string fieldName, List<string> errors)
        {
            if (value.TimeOfDay != TimeSpan.Zero)
                errors.Add(fieldName + " chi duoc chua ngay, khong chua gio.");
        }

        internal static bool IsSha256(string value)
        {
            return Sha256Pattern.IsMatch((value ?? string.Empty).Trim());
        }
    }
}
