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
    public sealed class PriceProfilePortfolio
    {
        public const int CurrentSchemaVersion = 1;
        private readonly IReadOnlyDictionary<MachineRateAudience, PriceProfile> byAudience;

        private PriceProfilePortfolio(IEnumerable<PriceProfile> profiles)
        {
            SchemaVersion = CurrentSchemaVersion;
            PriceProfile[] ordered = (profiles ?? Enumerable.Empty<PriceProfile>())
                .Where(profile => profile != null)
                .OrderBy(profile => profile.LaborAudience)
                .ToArray();
            Profiles = new ReadOnlyCollection<PriceProfile>(ordered);
            byAudience = new ReadOnlyDictionary<MachineRateAudience, PriceProfile>(
                ordered.ToDictionary(profile => profile.LaborAudience));
        }

        public int SchemaVersion { get; }
        public IReadOnlyList<PriceProfile> Profiles { get; }

        public static PriceProfilePortfolio Create(IEnumerable<PriceProfile> profiles)
        {
            var portfolio = new PriceProfilePortfolio(profiles);
            PriceProfilePortfolioValidator.ValidateRequired(portfolio);
            return portfolio;
        }

        public bool TryGet(MachineRateAudience audience, out PriceProfile profile)
        {
            return byAudience.TryGetValue(audience, out profile);
        }

        public PriceProfile FindRequired(MachineRateAudience audience)
        {
            PriceProfile profile;
            if (!TryGet(audience, out profile))
                throw new KeyNotFoundException("Chua co ho so gia cho doi tuong luong " + audience + ".");
            return profile;
        }

        public PriceProfilePortfolio WithProfile(PriceProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            return Create(Profiles
                .Where(item => item.LaborAudience != profile.LaborAudience)
                .Concat(new[] { profile }));
        }
    }

    public sealed class PriceProfilePortfolioValidationResult
    {
        internal PriceProfilePortfolioValidationResult(IEnumerable<string> errors)
        {
            Errors = new ReadOnlyCollection<string>(errors.ToList());
        }

        public IReadOnlyList<string> Errors { get; }
        public bool IsValid => Errors.Count == 0;
    }

    public static class PriceProfilePortfolioValidator
    {
        public static PriceProfilePortfolioValidationResult Validate(PriceProfilePortfolio portfolio)
        {
            var errors = new List<string>();
            if (portfolio == null)
            {
                errors.Add("PriceProfilePortfolio khong duoc null.");
                return new PriceProfilePortfolioValidationResult(errors);
            }
            if (portfolio.Profiles.Count < 1 || portfolio.Profiles.Count > 2)
                errors.Add("PriceProfilePortfolio phai co tu mot den hai ho so gia.");
            if (portfolio.Profiles.Select(profile => profile.LaborAudience).Distinct().Count() !=
                portfolio.Profiles.Count)
            {
                errors.Add("Moi doi tuong luong chi duoc co mot ho so gia.");
            }
            foreach (PriceProfile profile in portfolio.Profiles)
            {
                PriceProfileValidationResult validation = PriceProfileValidator.Validate(profile);
                errors.AddRange(validation.Errors.Select(error => profile.ProfileId + ": " + error));
            }
            return new PriceProfilePortfolioValidationResult(errors);
        }

        public static void ValidateRequired(PriceProfilePortfolio portfolio)
        {
            PriceProfilePortfolioValidationResult validation = Validate(portfolio);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(portfolio));
        }
    }

    public static class PriceProfilePortfolioSerializer
    {
        private const string Magic = "TTBMVN_PRICE_PROFILE_PORTFOLIO";

        public static string Serialize(PriceProfilePortfolio portfolio)
        {
            PriceProfilePortfolioValidator.ValidateRequired(portfolio);
            var lines = new List<string>
            {
                Magic,
                "schema=" + portfolio.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                "profileCount=" + portfolio.Profiles.Count.ToString(CultureInfo.InvariantCulture)
            };
            for (int index = 0; index < portfolio.Profiles.Count; index++)
            {
                PriceProfile profile = portfolio.Profiles[index];
                string prefix = "profile." + index.ToString(CultureInfo.InvariantCulture) + ".";
                lines.Add(prefix + "audience=" + ((int)profile.LaborAudience).ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "payload=" + Encode(PriceProfileSerializer.Serialize(profile)));
            }
            string body = string.Join("\n", lines);
            return body + "\nchecksum=" + Hash(body);
        }

        public static PriceProfilePortfolio Deserialize(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
                throw new InvalidDataException("PriceProfilePortfolio payload trong.");
            string[] lines = payload.Replace("\r\n", "\n").Split('\n');
            if (lines.Length < 5 || !string.Equals(lines[0], Magic, StringComparison.Ordinal))
                throw new InvalidDataException("PriceProfilePortfolio magic khong hop le.");
            string checksumLine = lines[lines.Length - 1];
            if (!checksumLine.StartsWith("checksum=", StringComparison.Ordinal))
                throw new InvalidDataException("PriceProfilePortfolio thieu checksum.");
            string body = string.Join("\n", lines.Take(lines.Length - 1));
            if (!string.Equals(checksumLine.Substring(9), Hash(body), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("PriceProfilePortfolio checksum khong khop.");

            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string line in lines.Skip(1).Take(lines.Length - 2))
            {
                int separator = line.IndexOf('=');
                if (separator <= 0 || values.ContainsKey(line.Substring(0, separator)))
                    throw new InvalidDataException("PriceProfilePortfolio field khong hop le.");
                values.Add(line.Substring(0, separator), line.Substring(separator + 1));
            }
            if (Integer(values, "schema") != PriceProfilePortfolio.CurrentSchemaVersion)
                throw new InvalidDataException("PriceProfilePortfolio schema khong duoc ho tro.");
            int count = Integer(values, "profileCount");
            if (count < 1 || count > 2)
                throw new InvalidDataException("PriceProfilePortfolio profileCount khong hop le.");
            var profiles = new List<PriceProfile>();
            for (int index = 0; index < count; index++)
            {
                string prefix = "profile." + index.ToString(CultureInfo.InvariantCulture) + ".";
                MachineRateAudience audience = (MachineRateAudience)Integer(values, prefix + "audience");
                PriceProfile profile = PriceProfileSerializer.Deserialize(Decode(Required(values, prefix + "payload")));
                if (profile.LaborAudience != audience)
                    throw new InvalidDataException("PriceProfilePortfolio audience khong khop payload.");
                profiles.Add(profile);
            }
            return PriceProfilePortfolio.Create(profiles);
        }

        private static string Required(IDictionary<string, string> values, string key)
        {
            string value;
            if (!values.TryGetValue(key, out value))
                throw new InvalidDataException("PriceProfilePortfolio thieu field " + key + ".");
            return value;
        }

        private static int Integer(IDictionary<string, string> values, string key)
        {
            int parsed;
            if (!int.TryParse(Required(values, key), NumberStyles.None, CultureInfo.InvariantCulture, out parsed))
                throw new InvalidDataException("PriceProfilePortfolio " + key + " khong hop le.");
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
                throw new InvalidDataException("PriceProfilePortfolio base64 khong hop le.", ex);
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
