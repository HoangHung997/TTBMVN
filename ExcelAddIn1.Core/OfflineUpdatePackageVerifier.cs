using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAddIn1.Core
{
    public static class OfflineUpdatePackageVerifier
    {
        private const long MaximumArchiveBytes = 100L * 1024L * 1024L;
        private const long MaximumManifestBytes = 1024L * 1024L;
        private const long MaximumSignatureBytes = 16L * 1024L;
        private const long MaximumPayloadFileBytes = 64L * 1024L * 1024L;
        private const long MaximumPayloadTotalBytes = 96L * 1024L * 1024L;
        private const int MaximumArchiveEntries = 66;

        public static OfflineUpdateVerificationResult Verify(
            string archivePath,
            IEnumerable<OfflineUpdateTrustedKey> trustedKeys,
            Version currentAppVersion,
            IEnumerable<RegulationPackage> installedPackages)
        {
            string archive = ValidateArchivePath(archivePath);
            if (currentAppVersion == null)
                throw new ArgumentNullException(nameof(currentAppVersion));
            Dictionary<string, OfflineUpdateTrustedKey> trust = BuildTrustStore(trustedKeys);

            try
            {
                using (FileStream stream = new FileStream(
                    archive, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8))
                {
                    Dictionary<string, ZipArchiveEntry> entries = IndexEntries(zip);
                    byte[] manifestBytes = ReadEntryRequired(
                        entries,
                        OfflineUpdateLayout.ManifestFileName,
                        MaximumManifestBytes);
                    string manifestText = new UTF8Encoding(false, true).GetString(manifestBytes);
                    OfflineUpdateManifest manifest;
                    try
                    {
                        manifest = OfflineUpdateManifestSerializer.Deserialize(manifestText);
                    }
                    catch (Exception ex) when (ex is InvalidDataException || ex is ArgumentException)
                    {
                        throw new OfflineUpdateVerificationException(
                            OfflineUpdateVerificationFailure.InvalidManifest,
                            "Manifest goi update khong hop le.",
                            ex);
                    }

                    ValidateRequiredPayloadCatalog(manifest);
                    OfflineUpdateTrustedKey key;
                    if (!trust.TryGetValue(manifest.SigningKeyId, out key))
                    {
                        throw new OfflineUpdateVerificationException(
                            OfflineUpdateVerificationFailure.UnknownSigningKey,
                            "Goi update duoc ky boi key khong duoc tin cay: " + manifest.SigningKeyId + ".");
                    }
                    byte[] signature = ReadEntryRequired(
                        entries,
                        OfflineUpdateLayout.SignatureFileName,
                        MaximumSignatureBytes);
                    if (!VerifySignature(manifestBytes, signature, key.PublicKey))
                    {
                        throw new OfflineUpdateVerificationException(
                            OfflineUpdateVerificationFailure.InvalidSignature,
                            "Chu ky so cua goi update khong hop le.");
                    }

                    ValidateArchiveCatalog(entries, manifest);
                    ValidatePayloadHashes(entries, manifest);
                    RegulationPackage package = ReadAndValidatePackage(entries, manifest);
                    OfflineUpdateDisposition disposition = EvaluateDisposition(
                        manifest,
                        currentAppVersion,
                        installedPackages,
                        out string reason);
                    return new OfflineUpdateVerificationResult(
                        archive,
                        ComputeFileChecksum(archive),
                        manifest,
                        package,
                        disposition,
                        reason);
                }
            }
            catch (OfflineUpdateVerificationException)
            {
                throw;
            }
            catch (Exception ex) when (
                ex is InvalidDataException || ex is IOException || ex is UnauthorizedAccessException ||
                ex is NotSupportedException || ex is CryptographicException || ex is DecoderFallbackException)
            {
                throw new OfflineUpdateVerificationException(
                    OfflineUpdateVerificationFailure.InvalidArchive,
                    "Khong doc duoc goi update offline.",
                    ex);
            }
        }

        public static string ExtractVerifiedBundle(
            OfflineUpdateVerificationResult verification,
            string destinationDirectory)
        {
            if (verification == null)
                throw new ArgumentNullException(nameof(verification));
            string currentChecksum = ComputeFileChecksum(ValidateArchivePath(verification.ArchivePath));
            if (!string.Equals(currentChecksum, verification.ArchiveChecksum, StringComparison.OrdinalIgnoreCase))
            {
                throw new OfflineUpdateVerificationException(
                    OfflineUpdateVerificationFailure.PayloadMismatch,
                    "Goi update da thay doi sau khi xac minh.");
            }
            string destination = Path.GetFullPath(destinationDirectory ?? string.Empty);
            if (Directory.Exists(destination) || File.Exists(destination))
                throw new IOException("Thu muc dich da ton tai: " + destination + ".");
            string parent = Path.GetDirectoryName(destination);
            if (string.IsNullOrWhiteSpace(parent))
                throw new ArgumentException("Thu muc dich khong hop le.", nameof(destinationDirectory));
            Directory.CreateDirectory(parent);
            string stage = destination + ".stage-" + Guid.NewGuid().ToString("N");
            try
            {
                Directory.CreateDirectory(stage);
                using (FileStream stream = new FileStream(
                    verification.ArchivePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8))
                {
                    Dictionary<string, ZipArchiveEntry> entries = IndexEntries(zip);
                    ValidateArchiveCatalog(entries, verification.Manifest);
                    ValidatePayloadHashes(entries, verification.Manifest);
                    ExtractPayload(entries, verification.Manifest, stage);
                }
                RegulationPackageBundle bundle = RegulationPackageBundleReader.Read(stage);
                EnsurePackageMatchesManifest(bundle.Package, verification.Manifest);
                Directory.Move(stage, destination);
                return destination;
            }
            catch
            {
                if (Directory.Exists(stage))
                    Directory.Delete(stage, true);
                throw;
            }
        }

        private static string ValidateArchivePath(string archivePath)
        {
            if (string.IsNullOrWhiteSpace(archivePath))
                throw new ArgumentException("Duong dan goi update la bat buoc.", nameof(archivePath));
            string fullPath = Path.GetFullPath(archivePath);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException("Khong tim thay goi update.", fullPath);
            var info = new FileInfo(fullPath);
            if (info.Length <= 0 || info.Length > MaximumArchiveBytes)
                throw new InvalidDataException("Kich thuoc goi update khong hop le.");
            if ((info.Attributes & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException("Goi update khong duoc la reparse point.");
            return fullPath;
        }

        private static Dictionary<string, OfflineUpdateTrustedKey> BuildTrustStore(
            IEnumerable<OfflineUpdateTrustedKey> trustedKeys)
        {
            var result = new Dictionary<string, OfflineUpdateTrustedKey>(StringComparer.Ordinal);
            foreach (OfflineUpdateTrustedKey key in trustedKeys ?? Enumerable.Empty<OfflineUpdateTrustedKey>())
            {
                if (key == null || result.ContainsKey(key.KeyId))
                    throw new ArgumentException("Trust store co key null hoac trung.", nameof(trustedKeys));
                result.Add(key.KeyId, key);
            }
            if (result.Count == 0)
                throw new ArgumentException("Trust store khong duoc rong.", nameof(trustedKeys));
            return result;
        }

        private static Dictionary<string, ZipArchiveEntry> IndexEntries(ZipArchive archive)
        {
            if (archive.Entries.Count < 3 || archive.Entries.Count > MaximumArchiveEntries)
                throw new InvalidDataException("So entry trong goi update khong hop le.");
            var result = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                string name = entry.FullName;
                if (!OfflineUpdateManifest.IsSafePayloadPath(name) ||
                    name.EndsWith("/", StringComparison.Ordinal) ||
                    result.ContainsKey(name))
                {
                    throw new InvalidDataException("Entry ZIP khong hop le hoac bi trung: " + name + ".");
                }
                result.Add(name, entry);
            }
            return result;
        }

        private static byte[] ReadEntryRequired(
            IDictionary<string, ZipArchiveEntry> entries,
            string path,
            long maximumBytes)
        {
            ZipArchiveEntry entry;
            if (!entries.TryGetValue(path, out entry))
                throw new InvalidDataException("Goi update thieu entry " + path + ".");
            if (entry.Length < 1 || entry.Length > maximumBytes)
                throw new InvalidDataException("Kich thuoc entry khong hop le: " + path + ".");
            using (Stream source = entry.Open())
            using (var target = new MemoryStream((int)entry.Length))
            {
                CopyBounded(source, target, maximumBytes);
                if (target.Length != entry.Length)
                    throw new InvalidDataException("Do dai entry ZIP khong khop: " + path + ".");
                return target.ToArray();
            }
        }

        private static void ValidateRequiredPayloadCatalog(OfflineUpdateManifest manifest)
        {
            string[] actual = manifest.Files.Select(file => file.Path)
                .OrderBy(path => path, StringComparer.Ordinal).ToArray();
            string[] expected = OfflineUpdateLayout.GetRequiredPayloadPaths().ToArray();
            if (!actual.SequenceEqual(expected, StringComparer.Ordinal))
            {
                throw new OfflineUpdateVerificationException(
                    OfflineUpdateVerificationFailure.InvalidManifest,
                    "Manifest update phai chua dung mot manifest package va sau module bat buoc.");
            }
        }

        private static void ValidateArchiveCatalog(
            IDictionary<string, ZipArchiveEntry> entries,
            OfflineUpdateManifest manifest)
        {
            var expected = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                OfflineUpdateLayout.ManifestFileName,
                OfflineUpdateLayout.SignatureFileName
            };
            foreach (OfflineUpdateFileManifest file in manifest.Files)
                expected.Add(file.Path);
            if (entries.Count != expected.Count || entries.Keys.Any(path => !expected.Contains(path)))
            {
                throw new OfflineUpdateVerificationException(
                    OfflineUpdateVerificationFailure.PayloadMismatch,
                    "Goi update co entry thua hoac thieu so voi manifest.");
            }
        }

        private static void ValidatePayloadHashes(
            IDictionary<string, ZipArchiveEntry> entries,
            OfflineUpdateManifest manifest)
        {
            long total = 0;
            foreach (OfflineUpdateFileManifest file in manifest.Files)
            {
                ZipArchiveEntry entry;
                if (!entries.TryGetValue(file.Path, out entry) ||
                    entry.Length != file.Length || entry.Length > MaximumPayloadFileBytes)
                {
                    throw new OfflineUpdateVerificationException(
                        OfflineUpdateVerificationFailure.PayloadMismatch,
                        "Do dai payload khong khop: " + file.Path + ".");
                }
                checked { total += entry.Length; }
                if (total > MaximumPayloadTotalBytes)
                    throw new InvalidDataException("Tong payload vuot qua gioi han.");
                using (Stream source = entry.Open())
                using (SHA256 sha256 = SHA256.Create())
                {
                    string checksum = ToHex(sha256.ComputeHash(source));
                    if (!string.Equals(checksum, file.Sha256, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new OfflineUpdateVerificationException(
                            OfflineUpdateVerificationFailure.PayloadMismatch,
                            "Checksum payload khong khop: " + file.Path + ".");
                    }
                }
            }
        }

        private static RegulationPackage ReadAndValidatePackage(
            IDictionary<string, ZipArchiveEntry> entries,
            OfflineUpdateManifest manifest)
        {
            string root = Path.Combine(Path.GetTempPath(), "TTBMVNUpdateVerify", Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(root);
                ExtractPayload(entries, manifest, root);
                RegulationPackageBundle bundle = RegulationPackageBundleReader.Read(root);
                EnsurePackageMatchesManifest(bundle.Package, manifest);
                return bundle.Package;
            }
            catch (OfflineUpdateVerificationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new OfflineUpdateVerificationException(
                    OfflineUpdateVerificationFailure.PackageMismatch,
                    "Payload khong tao thanh regulation package hop le.",
                    ex);
            }
            finally
            {
                if (Directory.Exists(root))
                    Directory.Delete(root, true);
            }
        }

        private static void ExtractPayload(
            IDictionary<string, ZipArchiveEntry> entries,
            OfflineUpdateManifest manifest,
            string destination)
        {
            string root = Path.GetFullPath(destination).TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            foreach (OfflineUpdateFileManifest file in manifest.Files)
            {
                string relative = file.Path.Substring(OfflineUpdateLayout.PayloadDirectoryName.Length + 1)
                    .Replace('/', Path.DirectorySeparatorChar);
                string target = Path.GetFullPath(Path.Combine(destination, relative));
                if (!target.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Payload thoat khoi thu muc dich.");
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                using (Stream source = entries[file.Path].Open())
                using (FileStream output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    CopyBounded(source, output, MaximumPayloadFileBytes);
            }
        }

        private static void EnsurePackageMatchesManifest(
            RegulationPackage package,
            OfflineUpdateManifest manifest)
        {
            if (!string.Equals(package.PackageId, manifest.PackageId, StringComparison.Ordinal) ||
                !string.Equals(package.DataVersion, manifest.PackageVersion, StringComparison.Ordinal) ||
                !string.Equals(package.PackageChecksum, manifest.PackageChecksum, StringComparison.OrdinalIgnoreCase))
            {
                throw new OfflineUpdateVerificationException(
                    OfflineUpdateVerificationFailure.PackageMismatch,
                    "Danh tinh regulation package khong khop manifest update.");
            }
        }

        private static OfflineUpdateDisposition EvaluateDisposition(
            OfflineUpdateManifest manifest,
            Version currentAppVersion,
            IEnumerable<RegulationPackage> installedPackages,
            out string reason)
        {
            var minimum = new Version(manifest.MinimumAppVersion);
            if (currentAppVersion < minimum)
            {
                reason = "Can app version toi thieu " + manifest.MinimumAppVersion + ".";
                return OfflineUpdateDisposition.AppUpgradeRequired;
            }
            List<RegulationPackage> samePackage = (installedPackages ?? Enumerable.Empty<RegulationPackage>())
                .Where(package => package != null &&
                    string.Equals(package.PackageId, manifest.PackageId, StringComparison.Ordinal))
                .ToList();
            RegulationPackage sameVersion = samePackage.FirstOrDefault(package =>
                string.Equals(package.DataVersion, manifest.PackageVersion, StringComparison.Ordinal));
            if (sameVersion != null)
            {
                if (string.Equals(
                    sameVersion.PackageChecksum,
                    manifest.PackageChecksum,
                    StringComparison.OrdinalIgnoreCase))
                {
                    reason = "Package cung version va checksum da duoc cai.";
                    return OfflineUpdateDisposition.AlreadyInstalled;
                }
                reason = "Version da ton tai voi checksum khac.";
                return OfflineUpdateDisposition.VersionConflict;
            }
            if (samePackage.Any(package =>
                RegulationDataVersionComparer.Compare(package.DataVersion, manifest.PackageVersion) > 0))
            {
                reason = "Khong cho cai package thap hon version dang co.";
                return OfflineUpdateDisposition.DowngradeRejected;
            }
            reason = "Goi update hop le va san sang cai dat.";
            return OfflineUpdateDisposition.Ready;
        }

        private static bool VerifySignature(byte[] data, byte[] signature, RSAParameters publicKey)
        {
            using (var rsa = new RSACryptoServiceProvider())
            {
                rsa.PersistKeyInCsp = false;
                rsa.ImportParameters(publicKey);
                return rsa.VerifyData(data, CryptoConfig.MapNameToOID("SHA256"), signature);
            }
        }

        private static void CopyBounded(Stream source, Stream target, long maximumBytes)
        {
            var buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;
                if (total > maximumBytes)
                    throw new InvalidDataException("Du lieu giai nen vuot qua gioi han.");
                target.Write(buffer, 0, read);
            }
        }

        private static string ComputeFileChecksum(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (SHA256 sha256 = SHA256.Create())
                return ToHex(sha256.ComputeHash(stream));
        }

        private static string ToHex(byte[] bytes)
        {
            var builder = new StringBuilder(bytes.Length * 2);
            foreach (byte value in bytes)
                builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
            return builder.ToString();
        }
    }
}
