using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace ExcelAddIn1.Core
{
    public static class OfflineUpdateLayout
    {
        public const string ManifestFileName = "update.ttbmanifest";
        public const string SignatureFileName = "update.signature";
        public const string PayloadDirectoryName = "payload";
        public const string FileExtension = ".ttbupdate";

        public static string GetPackageManifestPath()
        {
            return PayloadDirectoryName + "/" + RegulationPackageLayout.ManifestFileName;
        }

        public static string GetModulePath(RegulationModuleKind kind)
        {
            return PayloadDirectoryName + "/" +
                   RegulationPackageLayout.ModulesDirectoryName + "/" +
                   RegulationPackageLayout.GetModuleFileName(kind);
        }

        public static IReadOnlyList<string> GetRequiredPayloadPaths()
        {
            var paths = new List<string> { GetPackageManifestPath() };
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
                paths.Add(GetModulePath(kind));
            return new ReadOnlyCollection<string>(
                paths.OrderBy(path => path, StringComparer.Ordinal).ToList());
        }
    }

    public enum OfflineUpdateDisposition
    {
        Ready = 0,
        AlreadyInstalled = 1,
        DowngradeRejected = 2,
        VersionConflict = 3,
        AppUpgradeRequired = 4
    }

    public enum OfflineUpdateVerificationFailure
    {
        InvalidArchive = 1,
        InvalidManifest = 2,
        UnknownSigningKey = 3,
        InvalidSignature = 4,
        PayloadMismatch = 5,
        PackageMismatch = 6
    }

    public sealed class OfflineUpdateVerificationException : InvalidOperationException
    {
        public OfflineUpdateVerificationException(
            OfflineUpdateVerificationFailure failure,
            string message)
            : base(message)
        {
            Failure = failure;
        }

        public OfflineUpdateVerificationException(
            OfflineUpdateVerificationFailure failure,
            string message,
            Exception innerException)
            : base(message, innerException)
        {
            Failure = failure;
        }

        public OfflineUpdateVerificationFailure Failure { get; }
    }

    public sealed class OfflineUpdateFileManifest
    {
        public OfflineUpdateFileManifest(string path, long length, string sha256)
        {
            Path = path;
            Length = length;
            Sha256 = sha256;
        }

        public string Path { get; }
        public long Length { get; }
        public string Sha256 { get; }
    }

    public sealed class OfflineUpdateManifest
    {
        public const int CurrentSchemaVersion = 1;
        public const string SignatureAlgorithm = "RSA-SHA256-PKCS1";

        private OfflineUpdateManifest(
            int schemaVersion,
            string updateId,
            string updateVersion,
            string minimumAppVersion,
            string packageId,
            string packageVersion,
            string packageChecksum,
            string signingKeyId,
            string signatureAlgorithm,
            IEnumerable<OfflineUpdateFileManifest> files)
        {
            SchemaVersion = schemaVersion;
            UpdateId = updateId;
            UpdateVersion = updateVersion;
            MinimumAppVersion = minimumAppVersion;
            PackageId = packageId;
            PackageVersion = packageVersion;
            PackageChecksum = packageChecksum;
            SigningKeyId = signingKeyId;
            Algorithm = signatureAlgorithm;
            Files = new ReadOnlyCollection<OfflineUpdateFileManifest>(
                (files ?? Enumerable.Empty<OfflineUpdateFileManifest>()).ToList());
        }

        public int SchemaVersion { get; }
        public string UpdateId { get; }
        public string UpdateVersion { get; }
        public string MinimumAppVersion { get; }
        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public string SigningKeyId { get; }
        public string Algorithm { get; }
        public IReadOnlyList<OfflineUpdateFileManifest> Files { get; }

        public static OfflineUpdateManifest Create(
            string updateId,
            string updateVersion,
            string minimumAppVersion,
            string packageId,
            string packageVersion,
            string packageChecksum,
            string signingKeyId,
            IEnumerable<OfflineUpdateFileManifest> files)
        {
            return CreateCore(
                CurrentSchemaVersion,
                updateId,
                updateVersion,
                minimumAppVersion,
                packageId,
                packageVersion,
                packageChecksum,
                signingKeyId,
                SignatureAlgorithm,
                files);
        }

        internal static OfflineUpdateManifest Rehydrate(
            int schemaVersion,
            string updateId,
            string updateVersion,
            string minimumAppVersion,
            string packageId,
            string packageVersion,
            string packageChecksum,
            string signingKeyId,
            string signatureAlgorithm,
            IEnumerable<OfflineUpdateFileManifest> files)
        {
            return CreateCore(
                schemaVersion,
                updateId,
                updateVersion,
                minimumAppVersion,
                packageId,
                packageVersion,
                packageChecksum,
                signingKeyId,
                signatureAlgorithm,
                files);
        }

        private static OfflineUpdateManifest CreateCore(
            int schemaVersion,
            string updateId,
            string updateVersion,
            string minimumAppVersion,
            string packageId,
            string packageVersion,
            string packageChecksum,
            string signingKeyId,
            string signatureAlgorithm,
            IEnumerable<OfflineUpdateFileManifest> files)
        {
            if (schemaVersion != CurrentSchemaVersion)
                throw new ArgumentException("Schema goi update khong duoc ho tro.", nameof(schemaVersion));
            string normalizedUpdateId = RequireId(updateId, nameof(updateId));
            string normalizedUpdateVersion = RequireDataVersion(updateVersion, nameof(updateVersion));
            string normalizedMinimumAppVersion = NormalizeAppVersion(
                minimumAppVersion,
                nameof(minimumAppVersion));
            string normalizedPackageId = RequireId(packageId, nameof(packageId));
            string normalizedPackageVersion = RequireDataVersion(packageVersion, nameof(packageVersion));
            string normalizedChecksum = (packageChecksum ?? string.Empty).Trim().ToUpperInvariant();
            if (!RegulationPackageValidator.IsSha256(normalizedChecksum))
                throw new ArgumentException("Package checksum khong hop le.", nameof(packageChecksum));
            string normalizedKeyId = RequireId(signingKeyId, nameof(signingKeyId));
            if (!string.Equals(signatureAlgorithm, SignatureAlgorithm, StringComparison.Ordinal))
                throw new ArgumentException("Thuat toan chu ky khong duoc ho tro.", nameof(signatureAlgorithm));

            List<OfflineUpdateFileManifest> normalizedFiles = NormalizeFiles(files);
            return new OfflineUpdateManifest(
                schemaVersion,
                normalizedUpdateId,
                normalizedUpdateVersion,
                normalizedMinimumAppVersion,
                normalizedPackageId,
                normalizedPackageVersion,
                normalizedChecksum,
                normalizedKeyId,
                SignatureAlgorithm,
                normalizedFiles);
        }

        internal static string NormalizeAppVersion(string value, string parameterName)
        {
            string normalized = (value ?? string.Empty).Trim();
            Version version;
            if (!Regex.IsMatch(normalized, "^[0-9]+\\.[0-9]+\\.[0-9]+\\.[0-9]+$") ||
                !Version.TryParse(normalized, out version))
            {
                throw new ArgumentException(
                    "App version phai dung dang major.minor.build.revision.",
                    parameterName);
            }
            return version.ToString(4);
        }

        internal static bool IsSafePayloadPath(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length > 240)
                return false;
            if (value[0] == '/' || value.Contains("\\") || value.Contains(":") ||
                value.Contains("//"))
            {
                return false;
            }
            string[] segments = value.Split('/');
            return segments.All(segment =>
                segment.Length > 0 && segment != "." && segment != ".." &&
                Regex.IsMatch(segment, "^[A-Za-z0-9._-]+$"));
        }

        private static List<OfflineUpdateFileManifest> NormalizeFiles(
            IEnumerable<OfflineUpdateFileManifest> files)
        {
            var result = new List<OfflineUpdateFileManifest>();
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfflineUpdateFileManifest file in files ?? Enumerable.Empty<OfflineUpdateFileManifest>())
            {
                if (file == null)
                    throw new ArgumentException("File manifest khong duoc null.", nameof(files));
                string path = (file.Path ?? string.Empty).Trim();
                if (!IsSafePayloadPath(path) ||
                    !path.StartsWith(OfflineUpdateLayout.PayloadDirectoryName + "/", StringComparison.Ordinal))
                {
                    throw new ArgumentException("Duong dan payload khong hop le: " + path + ".", nameof(files));
                }
                if (!paths.Add(path))
                    throw new ArgumentException("Trung duong dan payload: " + path + ".", nameof(files));
                if (file.Length < 0)
                    throw new ArgumentException("Do dai payload khong duoc am.", nameof(files));
                string checksum = (file.Sha256 ?? string.Empty).Trim().ToUpperInvariant();
                if (!RegulationPackageValidator.IsSha256(checksum))
                    throw new ArgumentException("Checksum payload khong hop le: " + path + ".", nameof(files));
                result.Add(new OfflineUpdateFileManifest(path, file.Length, checksum));
            }
            if (result.Count == 0)
                throw new ArgumentException("Goi update phai co payload.", nameof(files));
            return result.OrderBy(file => file.Path, StringComparer.Ordinal).ToList();
        }

        private static string RequireId(string value, string parameterName)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (!Regex.IsMatch(normalized, "^[A-Za-z0-9][A-Za-z0-9._-]{2,159}$"))
                throw new ArgumentException("ID khong hop le.", parameterName);
            return normalized;
        }

        private static string RequireDataVersion(string value, string parameterName)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (!RegulationPackageValidator.IsDataVersion(normalized))
                throw new ArgumentException("Data version khong hop le.", parameterName);
            return normalized;
        }
    }

    public sealed class OfflineUpdateTrustedKey
    {
        public OfflineUpdateTrustedKey(string keyId, RSAParameters publicKey)
        {
            string normalized = (keyId ?? string.Empty).Trim();
            if (!Regex.IsMatch(normalized, "^[A-Za-z0-9][A-Za-z0-9._-]{2,159}$"))
                throw new ArgumentException("Signing key ID khong hop le.", nameof(keyId));
            if (publicKey.Modulus == null || publicKey.Modulus.Length < 256 ||
                publicKey.Exponent == null || publicKey.Exponent.Length == 0 ||
                publicKey.D != null)
            {
                throw new ArgumentException("Public key RSA khong hop le.", nameof(publicKey));
            }
            KeyId = normalized;
            PublicKey = publicKey;
        }

        public string KeyId { get; }
        public RSAParameters PublicKey { get; }
    }

    public sealed class OfflineUpdateVerificationResult
    {
        internal OfflineUpdateVerificationResult(
            string archivePath,
            string archiveChecksum,
            OfflineUpdateManifest manifest,
            RegulationPackage package,
            OfflineUpdateDisposition disposition,
            string dispositionReason)
        {
            ArchivePath = archivePath;
            ArchiveChecksum = archiveChecksum;
            Manifest = manifest;
            Package = package;
            Disposition = disposition;
            DispositionReason = dispositionReason;
        }

        public string ArchivePath { get; }
        public string ArchiveChecksum { get; }
        public OfflineUpdateManifest Manifest { get; }
        public RegulationPackage Package { get; }
        public OfflineUpdateDisposition Disposition { get; }
        public string DispositionReason { get; }
        public bool CanInstall => Disposition == OfflineUpdateDisposition.Ready;
    }

    public static class RegulationDataVersionComparer
    {
        public static int Compare(string left, string right)
        {
            SemanticVersion a = SemanticVersion.Parse(left);
            SemanticVersion b = SemanticVersion.Parse(right);
            int result = a.Major.CompareTo(b.Major);
            if (result != 0) return result;
            result = a.Minor.CompareTo(b.Minor);
            if (result != 0) return result;
            result = a.Patch.CompareTo(b.Patch);
            if (result != 0) return result;
            if (a.PreRelease.Count == 0 && b.PreRelease.Count == 0) return 0;
            if (a.PreRelease.Count == 0) return 1;
            if (b.PreRelease.Count == 0) return -1;

            int count = Math.Min(a.PreRelease.Count, b.PreRelease.Count);
            for (int index = 0; index < count; index++)
            {
                string ai = a.PreRelease[index];
                string bi = b.PreRelease[index];
                long an;
                long bn;
                bool aNumeric = long.TryParse(ai, NumberStyles.None, CultureInfo.InvariantCulture, out an);
                bool bNumeric = long.TryParse(bi, NumberStyles.None, CultureInfo.InvariantCulture, out bn);
                if (aNumeric && bNumeric)
                    result = an.CompareTo(bn);
                else if (aNumeric)
                    result = -1;
                else if (bNumeric)
                    result = 1;
                else
                    result = string.CompareOrdinal(ai, bi);
                if (result != 0) return result;
            }
            return a.PreRelease.Count.CompareTo(b.PreRelease.Count);
        }

        private sealed class SemanticVersion
        {
            private SemanticVersion(long major, long minor, long patch, IList<string> preRelease)
            {
                Major = major;
                Minor = minor;
                Patch = patch;
                PreRelease = preRelease;
            }

            internal long Major { get; }
            internal long Minor { get; }
            internal long Patch { get; }
            internal IList<string> PreRelease { get; }

            internal static SemanticVersion Parse(string value)
            {
                string normalized = (value ?? string.Empty).Trim();
                if (!RegulationPackageValidator.IsDataVersion(normalized))
                    throw new ArgumentException("Data version khong hop le.", nameof(value));
                string[] releaseParts = normalized.Split(new[] { '-' }, 2);
                string[] numbers = releaseParts[0].Split('.');
                long major;
                long minor;
                long patch;
                if (!long.TryParse(numbers[0], NumberStyles.None, CultureInfo.InvariantCulture, out major) ||
                    !long.TryParse(numbers[1], NumberStyles.None, CultureInfo.InvariantCulture, out minor) ||
                    !long.TryParse(numbers[2], NumberStyles.None, CultureInfo.InvariantCulture, out patch))
                {
                    throw new ArgumentException("Data version vuot qua gioi han so.", nameof(value));
                }
                IList<string> prerelease = releaseParts.Length == 1
                    ? new string[0]
                    : releaseParts[1].Split('.');
                return new SemanticVersion(major, minor, patch, prerelease);
            }
        }
    }
}
