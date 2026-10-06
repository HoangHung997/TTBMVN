using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAddIn1.Core
{
    public static class RegulationPackageSerializer
    {
        private const string Magic = "TTBMVN_REGULATION_PACKAGE";
        private const string DateFormat = "yyyy-MM-dd";
        private const int MaximumSources = 1000;
        private const int MaximumModules = 64;

        public static string Serialize(RegulationPackage package)
        {
            RegulationPackageValidationResult validation = RegulationPackageValidator.Validate(package);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(package));

            return BuildCanonicalBody(package) + "\nchecksum=" + package.PackageChecksum.ToUpperInvariant();
        }

        public static RegulationPackage Deserialize(string payload)
        {
            Dictionary<string, string> values = Parse(payload);
            int schemaVersion = ParseInt(GetRequired(values, "schema"), "schema", 1, int.MaxValue);
            if (schemaVersion != RegulationPackage.CurrentSchemaVersion)
                throw new InvalidDataException("Regulation package schema khong duoc ho tro: " + schemaVersion + ".");

            int sourceCount = ParseInt(
                GetRequired(values, "sourceCount"),
                "sourceCount",
                0,
                MaximumSources);
            int moduleCount = ParseInt(
                GetRequired(values, "moduleCount"),
                "moduleCount",
                0,
                MaximumModules);

            var sources = new List<RegulationPackageSourceDocument>();
            for (int index = 0; index < sourceCount; index++)
            {
                string prefix = "source." + index + ".";
                sources.Add(new RegulationPackageSourceDocument(
                    Decode(GetRequired(values, prefix + "documentId"), prefix + "documentId"),
                    Decode(GetRequired(values, prefix + "title"), prefix + "title"),
                    Decode(GetRequired(values, prefix + "publisher"), prefix + "publisher"),
                    ParseDate(GetRequired(values, prefix + "issuedDate"), prefix + "issuedDate", true).Value,
                    ParseDate(GetRequired(values, prefix + "effectiveFrom"), prefix + "effectiveFrom", true).Value,
                    ParseDate(GetRequired(values, prefix + "effectiveTo"), prefix + "effectiveTo", false),
                    Decode(GetRequired(values, prefix + "officialUri"), prefix + "officialUri"),
                    GetRequired(values, prefix + "contentChecksum")));
            }

            var modules = new List<RegulationPackageModuleManifest>();
            for (int index = 0; index < moduleCount; index++)
            {
                string prefix = "module." + index + ".";
                int kindValue = ParseInt(
                    GetRequired(values, prefix + "kind"),
                    prefix + "kind",
                    int.MinValue,
                    int.MaxValue);
                modules.Add(new RegulationPackageModuleManifest(
                    (RegulationModuleKind)kindValue,
                    Decode(GetRequired(values, prefix + "moduleId"), prefix + "moduleId"),
                    ParseInt(GetRequired(values, prefix + "schemaVersion"), prefix + "schemaVersion", 1, int.MaxValue),
                    Decode(GetRequired(values, prefix + "dataVersion"), prefix + "dataVersion"),
                    ParseInt(GetRequired(values, prefix + "recordCount"), prefix + "recordCount", 0, int.MaxValue),
                    GetRequired(values, prefix + "contentChecksum")));
            }

            int statusValue = ParseInt(
                GetRequired(values, "status"),
                "status",
                int.MinValue,
                int.MaxValue);
            RegulationPackage package = RegulationPackage.Rehydrate(
                schemaVersion,
                Decode(GetRequired(values, "packageId"), "packageId"),
                Decode(GetRequired(values, "dataVersion"), "dataVersion"),
                ParseDate(GetRequired(values, "effectiveFrom"), "effectiveFrom", true).Value,
                ParseDate(GetRequired(values, "effectiveTo"), "effectiveTo", false),
                (RegulationPackageStatus)statusValue,
                Decode(GetRequired(values, "transitionNote"), "transitionNote"),
                sources,
                modules,
                GetRequired(values, "checksum"));

            int expectedFields = 10 + (sourceCount * 8) + (moduleCount * 6);
            if (values.Count != expectedFields)
                throw new InvalidDataException("Regulation package co field thua hoac count khong khop.");

            RegulationPackageValidationResult validation = RegulationPackageValidator.Validate(package);
            if (!validation.IsValid)
                throw new InvalidDataException(string.Join(" ", validation.Errors));
            return package;
        }

        public static string ComputePackageChecksum(RegulationPackage package)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));
            return ComputeSha256(BuildCanonicalBody(package));
        }

        public static string ComputeSha256(string content)
        {
            if (content == null)
                throw new ArgumentNullException(nameof(content));
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(content));
                var builder = new StringBuilder(hash.Length * 2);
                foreach (byte value in hash)
                    builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
                return builder.ToString();
            }
        }

        private static string BuildCanonicalBody(RegulationPackage package)
        {
            RegulationPackageSourceDocument[] sources = package.Sources
                .OrderBy(source => source?.DocumentId ?? string.Empty, StringComparer.Ordinal)
                .ToArray();
            RegulationPackageModuleManifest[] modules = package.Modules
                .OrderBy(module => module == null ? int.MinValue : (int)module.Kind)
                .ThenBy(module => module?.ModuleId ?? string.Empty, StringComparer.Ordinal)
                .ToArray();
            var lines = new List<string>
            {
                Magic,
                "schema=" + package.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                "packageId=" + Encode(package.PackageId?.Trim() ?? string.Empty),
                "dataVersion=" + Encode(package.DataVersion?.Trim() ?? string.Empty),
                "effectiveFrom=" + FormatDate(package.EffectiveFrom),
                "effectiveTo=" + FormatDate(package.EffectiveTo),
                "status=" + ((int)package.Status).ToString(CultureInfo.InvariantCulture),
                "transitionNote=" + Encode(package.TransitionNote ?? string.Empty),
                "sourceCount=" + sources.Length.ToString(CultureInfo.InvariantCulture),
                "moduleCount=" + modules.Length.ToString(CultureInfo.InvariantCulture)
            };

            for (int index = 0; index < sources.Length; index++)
            {
                RegulationPackageSourceDocument source = sources[index];
                string prefix = "source." + index + ".";
                lines.Add(prefix + "documentId=" + Encode(source?.DocumentId?.Trim() ?? string.Empty));
                lines.Add(prefix + "title=" + Encode(source?.Title?.Trim() ?? string.Empty));
                lines.Add(prefix + "publisher=" + Encode(source?.Publisher?.Trim() ?? string.Empty));
                lines.Add(prefix + "issuedDate=" + FormatDate(source?.IssuedDate));
                lines.Add(prefix + "effectiveFrom=" + FormatDate(source?.EffectiveFrom));
                lines.Add(prefix + "effectiveTo=" + FormatDate(source?.EffectiveTo));
                lines.Add(prefix + "officialUri=" + Encode(source?.OfficialUri?.Trim() ?? string.Empty));
                lines.Add(prefix + "contentChecksum=" + (source?.ContentChecksum?.Trim().ToUpperInvariant() ?? string.Empty));
            }

            for (int index = 0; index < modules.Length; index++)
            {
                RegulationPackageModuleManifest module = modules[index];
                string prefix = "module." + index + ".";
                lines.Add(prefix + "kind=" + (module == null
                    ? string.Empty
                    : ((int)module.Kind).ToString(CultureInfo.InvariantCulture)));
                lines.Add(prefix + "moduleId=" + Encode(module?.ModuleId?.Trim() ?? string.Empty));
                lines.Add(prefix + "schemaVersion=" + (module == null
                    ? string.Empty
                    : module.SchemaVersion.ToString(CultureInfo.InvariantCulture)));
                lines.Add(prefix + "dataVersion=" + Encode(module?.DataVersion?.Trim() ?? string.Empty));
                lines.Add(prefix + "recordCount=" + (module == null
                    ? string.Empty
                    : module.RecordCount.ToString(CultureInfo.InvariantCulture)));
                lines.Add(prefix + "contentChecksum=" + (module?.ContentChecksum?.Trim().ToUpperInvariant() ?? string.Empty));
            }

            return string.Join("\n", lines);
        }

        private static Dictionary<string, string> Parse(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
                throw new InvalidDataException("Regulation package payload trong.");
            string normalized = payload.Replace("\r\n", "\n").Replace('\r', '\n');
            string[] lines = normalized.Split(new[] { '\n' }, StringSplitOptions.None);
            if (lines.Length == 0 || !string.Equals(lines[0], Magic, StringComparison.Ordinal))
                throw new InvalidDataException("Regulation package payload khong dung dinh dang.");

            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int index = 1; index < lines.Length; index++)
            {
                int separator = lines[index].IndexOf('=');
                if (separator <= 0)
                    throw new InvalidDataException("Regulation package co dong khong hop le tai vi tri " + (index + 1) + ".");
                string key = lines[index].Substring(0, separator);
                if (values.ContainsKey(key))
                    throw new InvalidDataException("Regulation package bi trung field: " + key + ".");
                values.Add(key, lines[index].Substring(separator + 1));
            }
            return values;
        }

        private static string GetRequired(Dictionary<string, string> values, string key)
        {
            string value;
            if (!values.TryGetValue(key, out value))
                throw new InvalidDataException("Regulation package thieu field: " + key + ".");
            return value;
        }

        private static int ParseInt(string value, string fieldName, int minimum, int maximum)
        {
            int parsed;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ||
                parsed < minimum || parsed > maximum)
            {
                throw new InvalidDataException("Regulation package field " + fieldName + " khong hop le.");
            }
            return parsed;
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value, string fieldName)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value ?? string.Empty));
            }
            catch (FormatException ex)
            {
                throw new InvalidDataException(
                    "Regulation package field " + fieldName + " khong phai Base64 hop le.",
                    ex);
            }
        }

        private static string FormatDate(DateTime value)
        {
            return value.ToString(DateFormat, CultureInfo.InvariantCulture);
        }

        private static string FormatDate(DateTime? value)
        {
            return value.HasValue ? FormatDate(value.Value) : string.Empty;
        }

        private static DateTime? ParseDate(string value, string fieldName, bool required)
        {
            if (string.IsNullOrEmpty(value))
            {
                if (required)
                    throw new InvalidDataException("Regulation package field " + fieldName + " la bat buoc.");
                return null;
            }
            DateTime parsed;
            if (!DateTime.TryParseExact(
                value,
                DateFormat,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out parsed))
            {
                throw new InvalidDataException(
                    "Regulation package field " + fieldName + " khong dung yyyy-MM-dd.");
            }
            return parsed;
        }
    }
}
