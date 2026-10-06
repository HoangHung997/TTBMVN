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
    public sealed class ProjectProfile
    {
        public const int CurrentSchemaVersion = 2;

        public int SchemaVersion { get; set; } = CurrentSchemaVersion;
        public string ProjectId { get; set; }
        public DateTime? PreparedDate { get; set; }
        public DateTime? ApprovalDate { get; set; }
        public DateTime? PriceDate { get; set; }
        public string RegulationPackageId { get; set; }
        public string RegulationPackageVersion { get; set; }
        public string RegulationPackageChecksum { get; set; }
        public string PriceProfileId { get; set; }
        public string OverrideSummary { get; set; }

        public ProjectProfile Clone()
        {
            return new ProjectProfile
            {
                SchemaVersion = SchemaVersion,
                ProjectId = ProjectId,
                PreparedDate = PreparedDate,
                ApprovalDate = ApprovalDate,
                PriceDate = PriceDate,
                RegulationPackageId = RegulationPackageId,
                RegulationPackageVersion = RegulationPackageVersion,
                RegulationPackageChecksum = RegulationPackageChecksum,
                PriceProfileId = PriceProfileId,
                OverrideSummary = OverrideSummary
            };
        }
    }

    public sealed class ProjectProfileValidationResult
    {
        internal ProjectProfileValidationResult(IEnumerable<string> errors)
        {
            Errors = new ReadOnlyCollection<string>((errors ?? Enumerable.Empty<string>()).ToList());
        }

        public bool IsValid => Errors.Count == 0;
        public IReadOnlyList<string> Errors { get; }
    }

    public static class ProjectProfileValidator
    {
        public static ProjectProfileValidationResult Validate(ProjectProfile profile)
        {
            var errors = new List<string>();
            if (profile == null)
            {
                errors.Add("Project profile khong duoc null.");
                return new ProjectProfileValidationResult(errors);
            }

            if (profile.SchemaVersion != ProjectProfile.CurrentSchemaVersion)
                errors.Add("Schema version khong duoc ho tro: " + profile.SchemaVersion + ".");

            ValidateRequiredId(profile.ProjectId, "ProjectId", 128, errors);
            ValidateRequiredId(profile.RegulationPackageId, "RegulationPackageId", 160, errors);
            string packageVersion = (profile.RegulationPackageVersion ?? string.Empty).Trim();
            string packageChecksum = (profile.RegulationPackageChecksum ?? string.Empty).Trim();
            if ((packageVersion.Length == 0) != (packageChecksum.Length == 0))
                errors.Add("RegulationPackageVersion va RegulationPackageChecksum phai cung co hoac cung trong.");
            if (packageVersion.Length > 0 && !RegulationPackageValidator.IsDataVersion(packageVersion))
                errors.Add("RegulationPackageVersion phai dung dang major.minor.patch.");
            if (packageChecksum.Length > 0 && !RegulationPackageValidator.IsSha256(packageChecksum))
                errors.Add("RegulationPackageChecksum phai la SHA-256 hex 64 ky tu.");
            ValidateRequiredId(profile.PriceProfileId, "PriceProfileId", 160, errors);

            ValidateDate(profile.PreparedDate, "PreparedDate", true, errors);
            ValidateDate(profile.ApprovalDate, "ApprovalDate", false, errors);
            ValidateDate(profile.PriceDate, "PriceDate", true, errors);

            if (profile.PreparedDate.HasValue && profile.ApprovalDate.HasValue &&
                profile.ApprovalDate.Value.Date < profile.PreparedDate.Value.Date)
            {
                errors.Add("ApprovalDate khong duoc truoc PreparedDate.");
            }

            if ((profile.OverrideSummary ?? string.Empty).Length > 4096)
                errors.Add("OverrideSummary khong duoc vuot qua 4096 ky tu.");

            return new ProjectProfileValidationResult(errors);
        }

        private static void ValidateRequiredId(string value, string fieldName, int maxLength, List<string> errors)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length == 0)
            {
                errors.Add(fieldName + " la bat buoc.");
                return;
            }

            if (normalized.Length > maxLength)
                errors.Add(fieldName + " khong duoc vuot qua " + maxLength + " ky tu.");
        }

        private static void ValidateDate(
            DateTime? value,
            string fieldName,
            bool required,
            List<string> errors)
        {
            if (!value.HasValue)
            {
                if (required)
                    errors.Add(fieldName + " la bat buoc.");
                return;
            }

            if (value.Value.TimeOfDay != TimeSpan.Zero)
                errors.Add(fieldName + " chi duoc chua ngay, khong chua gio.");
        }
    }

    public static class ProjectProfileSerializer
    {
        private const string Magic = "TTBMVN_PROJECT_PROFILE";
        private const string DateFormat = "yyyy-MM-dd";

        public static string Serialize(ProjectProfile profile)
        {
            ProjectProfileValidationResult validation = ProjectProfileValidator.Validate(profile);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(profile));

            var lines = new[]
            {
                Magic,
                "schema=" + profile.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                "projectId=" + Encode(profile.ProjectId.Trim()),
                "preparedDate=" + FormatDate(profile.PreparedDate),
                "approvalDate=" + FormatDate(profile.ApprovalDate),
                "priceDate=" + FormatDate(profile.PriceDate),
                "regulationPackageId=" + Encode(profile.RegulationPackageId.Trim()),
                "regulationPackageVersion=" + Encode((profile.RegulationPackageVersion ?? string.Empty).Trim()),
                "regulationPackageChecksum=" + (profile.RegulationPackageChecksum ?? string.Empty).Trim().ToUpperInvariant(),
                "priceProfileId=" + Encode(profile.PriceProfileId.Trim()),
                "overrideSummary=" + Encode(profile.OverrideSummary ?? string.Empty)
            };

            return string.Join("\n", lines);
        }

        public static ProjectProfile Deserialize(string payload)
        {
            Dictionary<string, string> values = Parse(payload);
            int schemaVersion = ParseSchemaVersion(values);
            ProjectProfile profile;

            switch (schemaVersion)
            {
                case 0:
                    EnsureFieldCount(values, 8, schemaVersion);
                    profile = ReadSchema0(values);
                    break;
                case 1:
                    EnsureFieldCount(values, 8, schemaVersion);
                    profile = ReadSchema1(values);
                    break;
                case ProjectProfile.CurrentSchemaVersion:
                    EnsureFieldCount(values, 10, schemaVersion);
                    profile = ReadSchema2(values);
                    break;
                default:
                    throw new InvalidDataException("Project profile schema khong duoc ho tro: " + schemaVersion + ".");
            }

            ProjectProfileValidationResult validation = ProjectProfileValidator.Validate(profile);
            if (!validation.IsValid)
                throw new InvalidDataException(string.Join(" ", validation.Errors));

            return profile;
        }

        public static string ComputeChecksum(string payload)
        {
            if (payload == null)
                throw new ArgumentNullException(nameof(payload));

            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(payload));
                var builder = new StringBuilder(hash.Length * 2);
                foreach (byte value in hash)
                    builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
                return builder.ToString();
            }
        }

        private static Dictionary<string, string> Parse(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
                throw new InvalidDataException("Project profile payload trong.");

            string normalized = payload.Replace("\r\n", "\n").Replace('\r', '\n');
            string[] lines = normalized.Split(new[] { '\n' }, StringSplitOptions.None);
            if (lines.Length == 0 || !string.Equals(lines[0], Magic, StringComparison.Ordinal))
                throw new InvalidDataException("Project profile payload khong dung dinh dang.");

            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int index = 1; index < lines.Length; index++)
            {
                string line = lines[index];
                int separator = line.IndexOf('=');
                if (separator <= 0)
                    throw new InvalidDataException("Project profile co dong khong hop le tai vi tri " + (index + 1) + ".");

                string key = line.Substring(0, separator);
                string value = line.Substring(separator + 1);
                if (values.ContainsKey(key))
                    throw new InvalidDataException("Project profile bi trung field: " + key + ".");
                values.Add(key, value);
            }

            return values;
        }

        private static int ParseSchemaVersion(Dictionary<string, string> values)
        {
            string text = GetRequired(values, "schema");
            int schemaVersion;
            if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out schemaVersion) || schemaVersion < 0)
                throw new InvalidDataException("Project profile schema khong hop le.");
            return schemaVersion;
        }

        private static void EnsureFieldCount(
            Dictionary<string, string> values,
            int expectedCount,
            int schemaVersion)
        {
            if (values.Count != expectedCount)
            {
                throw new InvalidDataException(
                    "Project profile schema " + schemaVersion + " co field thua hoac thieu.");
            }
        }

        private static ProjectProfile ReadSchema1(Dictionary<string, string> values)
        {
            return new ProjectProfile
            {
                SchemaVersion = ProjectProfile.CurrentSchemaVersion,
                ProjectId = Decode(GetRequired(values, "projectId"), "projectId"),
                PreparedDate = ParseDate(GetRequired(values, "preparedDate"), "preparedDate", true),
                ApprovalDate = ParseDate(GetRequired(values, "approvalDate"), "approvalDate", false),
                PriceDate = ParseDate(GetRequired(values, "priceDate"), "priceDate", true),
                RegulationPackageId = Decode(GetRequired(values, "regulationPackageId"), "regulationPackageId"),
                RegulationPackageVersion = string.Empty,
                RegulationPackageChecksum = string.Empty,
                PriceProfileId = Decode(GetRequired(values, "priceProfileId"), "priceProfileId"),
                OverrideSummary = Decode(GetRequired(values, "overrideSummary"), "overrideSummary")
            };
        }

        private static ProjectProfile ReadSchema2(Dictionary<string, string> values)
        {
            return new ProjectProfile
            {
                SchemaVersion = ProjectProfile.CurrentSchemaVersion,
                ProjectId = Decode(GetRequired(values, "projectId"), "projectId"),
                PreparedDate = ParseDate(GetRequired(values, "preparedDate"), "preparedDate", true),
                ApprovalDate = ParseDate(GetRequired(values, "approvalDate"), "approvalDate", false),
                PriceDate = ParseDate(GetRequired(values, "priceDate"), "priceDate", true),
                RegulationPackageId = Decode(GetRequired(values, "regulationPackageId"), "regulationPackageId"),
                RegulationPackageVersion = Decode(
                    GetRequired(values, "regulationPackageVersion"),
                    "regulationPackageVersion"),
                RegulationPackageChecksum = GetRequired(values, "regulationPackageChecksum"),
                PriceProfileId = Decode(GetRequired(values, "priceProfileId"), "priceProfileId"),
                OverrideSummary = Decode(GetRequired(values, "overrideSummary"), "overrideSummary")
            };
        }

        private static ProjectProfile ReadSchema0(Dictionary<string, string> values)
        {
            return new ProjectProfile
            {
                SchemaVersion = ProjectProfile.CurrentSchemaVersion,
                ProjectId = Decode(GetRequired(values, "project"), "project"),
                PreparedDate = ParseDate(GetRequired(values, "created"), "created", true),
                ApprovalDate = ParseDate(GetRequired(values, "approved"), "approved", false),
                PriceDate = ParseDate(GetRequired(values, "price"), "price", true),
                RegulationPackageId = Decode(GetRequired(values, "package"), "package"),
                RegulationPackageVersion = string.Empty,
                RegulationPackageChecksum = string.Empty,
                PriceProfileId = Decode(GetRequired(values, "priceProfile"), "priceProfile"),
                OverrideSummary = Decode(GetRequired(values, "overrides"), "overrides")
            };
        }

        private static string GetRequired(Dictionary<string, string> values, string key)
        {
            string value;
            if (!values.TryGetValue(key, out value))
                throw new InvalidDataException("Project profile thieu field: " + key + ".");
            return value;
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
                throw new InvalidDataException("Project profile field " + fieldName + " khong phai Base64 hop le.", ex);
            }
        }

        private static string FormatDate(DateTime? value)
        {
            return value.HasValue
                ? value.Value.ToString(DateFormat, CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private static DateTime? ParseDate(string value, string fieldName, bool required)
        {
            if (string.IsNullOrEmpty(value))
            {
                if (required)
                    throw new InvalidDataException("Project profile field " + fieldName + " la bat buoc.");
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
                throw new InvalidDataException("Project profile field " + fieldName + " khong dung yyyy-MM-dd.");
            }

            return parsed;
        }
    }
}
