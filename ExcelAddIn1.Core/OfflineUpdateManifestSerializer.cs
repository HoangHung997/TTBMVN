using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAddIn1.Core
{
    public static class OfflineUpdateManifestSerializer
    {
        private const string Header = "TTBMVN_OFFLINE_UPDATE";

        public static string Serialize(OfflineUpdateManifest manifest)
        {
            if (manifest == null)
                throw new ArgumentNullException(nameof(manifest));
            var lines = new List<string>
            {
                Header,
                "schema=" + manifest.SchemaVersion.ToString(CultureInfo.InvariantCulture),
                "updateId=" + Encode(manifest.UpdateId),
                "updateVersion=" + manifest.UpdateVersion,
                "minimumAppVersion=" + manifest.MinimumAppVersion,
                "packageId=" + Encode(manifest.PackageId),
                "packageVersion=" + manifest.PackageVersion,
                "packageChecksum=" + manifest.PackageChecksum,
                "signingKeyId=" + Encode(manifest.SigningKeyId),
                "algorithm=" + manifest.Algorithm,
                "fileCount=" + manifest.Files.Count.ToString(CultureInfo.InvariantCulture)
            };
            for (int index = 0; index < manifest.Files.Count; index++)
            {
                OfflineUpdateFileManifest file = manifest.Files[index];
                string prefix = "file." + index.ToString(CultureInfo.InvariantCulture) + ".";
                lines.Add(prefix + "path=" + file.Path);
                lines.Add(prefix + "length=" + file.Length.ToString(CultureInfo.InvariantCulture));
                lines.Add(prefix + "sha256=" + file.Sha256);
            }
            return string.Join("\n", lines) + "\n";
        }

        public static OfflineUpdateManifest Deserialize(string payload)
        {
            if (payload == null)
                throw new ArgumentNullException(nameof(payload));
            string normalized = payload.Replace("\r\n", "\n");
            if (!string.Equals(payload, normalized, StringComparison.Ordinal) ||
                !normalized.EndsWith("\n", StringComparison.Ordinal))
            {
                throw new InvalidDataException("Manifest update phai dung UTF-8/LF canonical.");
            }
            string[] lines = normalized.Split('\n');
            if (lines.Length < 12 || !string.Equals(lines[0], Header, StringComparison.Ordinal))
                throw new InvalidDataException("Header manifest update khong hop le.");

            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int index = 1; index < lines.Length - 1; index++)
            {
                int separator = lines[index].IndexOf('=');
                if (separator <= 0)
                    throw new InvalidDataException("Dong manifest update khong hop le: " + (index + 1) + ".");
                string key = lines[index].Substring(0, separator);
                string value = lines[index].Substring(separator + 1);
                if (values.ContainsKey(key))
                    throw new InvalidDataException("Trung property manifest update: " + key + ".");
                values.Add(key, value);
            }

            int schema = ParseInt(Get(values, "schema"), "schema");
            int fileCount = ParseInt(Get(values, "fileCount"), "fileCount");
            if (fileCount < 1 || fileCount > 64)
                throw new InvalidDataException("So file payload khong hop le.");
            var files = new List<OfflineUpdateFileManifest>();
            for (int index = 0; index < fileCount; index++)
            {
                string prefix = "file." + index.ToString(CultureInfo.InvariantCulture) + ".";
                long length = ParseLong(Get(values, prefix + "length"), prefix + "length");
                files.Add(new OfflineUpdateFileManifest(
                    Get(values, prefix + "path"),
                    length,
                    Get(values, prefix + "sha256")));
            }

            var expected = new HashSet<string>(StringComparer.Ordinal)
            {
                "schema", "updateId", "updateVersion", "minimumAppVersion", "packageId",
                "packageVersion", "packageChecksum", "signingKeyId", "algorithm", "fileCount"
            };
            for (int index = 0; index < fileCount; index++)
            {
                string prefix = "file." + index.ToString(CultureInfo.InvariantCulture) + ".";
                expected.Add(prefix + "path");
                expected.Add(prefix + "length");
                expected.Add(prefix + "sha256");
            }
            string[] unexpected = values.Keys.Where(key => !expected.Contains(key)).ToArray();
            string[] missing = expected.Where(key => !values.ContainsKey(key)).ToArray();
            if (unexpected.Length > 0 || missing.Length > 0 || values.Count != expected.Count)
                throw new InvalidDataException("Property manifest update thua hoac thieu.");

            OfflineUpdateManifest manifest;
            try
            {
                manifest = OfflineUpdateManifest.Rehydrate(
                    schema,
                    Decode(Get(values, "updateId")),
                    Get(values, "updateVersion"),
                    Get(values, "minimumAppVersion"),
                    Decode(Get(values, "packageId")),
                    Get(values, "packageVersion"),
                    Get(values, "packageChecksum"),
                    Decode(Get(values, "signingKeyId")),
                    Get(values, "algorithm"),
                    files);
            }
            catch (ArgumentException ex)
            {
                throw new InvalidDataException("Manifest update khong hop le: " + ex.Message, ex);
            }
            if (!string.Equals(Serialize(manifest), payload, StringComparison.Ordinal))
                throw new InvalidDataException("Manifest update khong o dang canonical.");
            return manifest;
        }

        private static int ParseInt(string value, string field)
        {
            int result;
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out result))
                throw new InvalidDataException(field + " khong hop le.");
            return result;
        }

        private static long ParseLong(string value, string field)
        {
            long result;
            if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out result) || result < 0)
                throw new InvalidDataException(field + " khong hop le.");
            return result;
        }

        private static string Get(IDictionary<string, string> values, string key)
        {
            string value;
            if (!values.TryGetValue(key, out value))
                throw new InvalidDataException("Thieu property manifest update: " + key + ".");
            return value;
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(new UTF8Encoding(false).GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value)
        {
            try
            {
                byte[] bytes = Convert.FromBase64String(value);
                string decoded = new UTF8Encoding(false, true).GetString(bytes);
                if (!string.Equals(Encode(decoded), value, StringComparison.Ordinal))
                    throw new FormatException();
                return decoded;
            }
            catch (Exception ex) when (ex is FormatException || ex is DecoderFallbackException)
            {
                throw new InvalidDataException("Gia tri Base64 trong manifest update khong hop le.", ex);
            }
        }
    }

    public static class OfflineUpdatePublicKeySerializer
    {
        private const string Header = "TTBMVN_OFFLINE_PUBLIC_KEY";

        public static string Serialize(OfflineUpdateTrustedKey key)
        {
            if (key == null)
                throw new ArgumentNullException(nameof(key));
            return string.Join("\n", new[]
            {
                Header,
                "schema=1",
                "keyId=" + Convert.ToBase64String(Encoding.UTF8.GetBytes(key.KeyId)),
                "algorithm=" + OfflineUpdateManifest.SignatureAlgorithm,
                "modulus=" + Convert.ToBase64String(key.PublicKey.Modulus),
                "exponent=" + Convert.ToBase64String(key.PublicKey.Exponent),
                string.Empty
            });
        }

        public static OfflineUpdateTrustedKey Deserialize(string payload)
        {
            if (payload == null)
                throw new ArgumentNullException(nameof(payload));
            string[] lines = payload.Split('\n');
            if (lines.Length != 7 || lines[0] != Header || lines[1] != "schema=1" ||
                lines[3] != "algorithm=" + OfflineUpdateManifest.SignatureAlgorithm ||
                lines[6].Length != 0)
            {
                throw new InvalidDataException("Public key update khong dung dinh dang canonical.");
            }
            try
            {
                string keyId = Encoding.UTF8.GetString(ReadProperty(lines[2], "keyId"));
                var parameters = new RSAParameters
                {
                    Modulus = ReadProperty(lines[4], "modulus"),
                    Exponent = ReadProperty(lines[5], "exponent")
                };
                var key = new OfflineUpdateTrustedKey(keyId, parameters);
                if (!string.Equals(Serialize(key), payload, StringComparison.Ordinal))
                    throw new InvalidDataException("Public key update khong o dang canonical.");
                return key;
            }
            catch (FormatException ex)
            {
                throw new InvalidDataException("Public key update co Base64 khong hop le.", ex);
            }
        }

        private static byte[] ReadProperty(string line, string key)
        {
            string prefix = key + "=";
            if (!line.StartsWith(prefix, StringComparison.Ordinal))
                throw new InvalidDataException("Thieu property public key: " + key + ".");
            return Convert.FromBase64String(line.Substring(prefix.Length));
        }
    }
}
