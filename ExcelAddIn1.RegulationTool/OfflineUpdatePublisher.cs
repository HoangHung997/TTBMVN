using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using ExcelAddIn1.Core;

namespace ExcelAddIn1.RegulationTool
{
    internal static class OfflineUpdatePublisher
    {
        private const int RsaProviderType = 24;
        private static readonly DateTimeOffset DeterministicZipTimestamp =
            new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);

        internal static OfflineUpdateTrustedKey CreateKey(
            string containerName,
            string publicKeyPath,
            string keyId)
        {
            ValidateContainerName(containerName);
            string output = PrepareNewOutputFile(publicKeyPath);
            using (RSACryptoServiceProvider rsa = OpenContainer(containerName, create: true))
            {
                var trustedKey = new OfflineUpdateTrustedKey(keyId, rsa.ExportParameters(false));
                File.WriteAllText(
                    output,
                    OfflineUpdatePublicKeySerializer.Serialize(trustedKey),
                    new UTF8Encoding(false));
                return trustedKey;
            }
        }

        internal static OfflineUpdateManifest Build(
            string bundleDirectory,
            string outputFile,
            string minimumAppVersion,
            string containerName,
            string publicKeyPath)
        {
            ValidateContainerName(containerName);
            RegulationPackageBundle bundle = RegulationPackageBundleReader.Read(bundleDirectory);
            OfflineUpdateTrustedKey trustedKey = OfflineUpdatePublicKeySerializer.Deserialize(
                File.ReadAllText(publicKeyPath, new UTF8Encoding(false, true)));
            var sourceFiles = GetPayloadSourceFiles(bundle.Directory);
            var fileManifests = sourceFiles
                .Select(item => new OfflineUpdateFileManifest(
                    item.Key,
                    new FileInfo(item.Value).Length,
                    ComputeFileChecksum(item.Value)))
                .ToArray();
            OfflineUpdateManifest manifest = OfflineUpdateManifest.Create(
                bundle.Package.PackageId + "-" + bundle.Package.DataVersion,
                bundle.Package.DataVersion,
                minimumAppVersion,
                bundle.Package.PackageId,
                bundle.Package.DataVersion,
                bundle.Package.PackageChecksum,
                trustedKey.KeyId,
                fileManifests);
            byte[] manifestBytes = new UTF8Encoding(false).GetBytes(
                OfflineUpdateManifestSerializer.Serialize(manifest));

            byte[] signature;
            using (RSACryptoServiceProvider rsa = OpenContainer(containerName, create: false))
            {
                EnsureSamePublicKey(rsa.ExportParameters(false), trustedKey.PublicKey);
                signature = rsa.SignData(manifestBytes, CryptoConfig.MapNameToOID("SHA256"));
            }

            string output = Path.GetFullPath(outputFile ?? string.Empty);
            if (!output.EndsWith(OfflineUpdateLayout.FileExtension, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("File update phai co duoi " + OfflineUpdateLayout.FileExtension + ".", nameof(outputFile));
            if (File.Exists(output) || Directory.Exists(output))
                throw new IOException("File update dich da ton tai: " + output + ".");
            string parent = Path.GetDirectoryName(output);
            if (string.IsNullOrWhiteSpace(parent))
                throw new ArgumentException("Thu muc dich khong hop le.", nameof(outputFile));
            Directory.CreateDirectory(parent);
            string stage = output + ".stage-" + Guid.NewGuid().ToString("N");
            try
            {
                using (FileStream stream = new FileStream(stage, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, false, Encoding.UTF8))
                {
                    WriteEntry(zip, OfflineUpdateLayout.ManifestFileName, manifestBytes);
                    WriteEntry(zip, OfflineUpdateLayout.SignatureFileName, signature);
                    foreach (KeyValuePair<string, string> item in sourceFiles.OrderBy(item => item.Key, StringComparer.Ordinal))
                        WriteFileEntry(zip, item.Key, item.Value);
                }
                File.Move(stage, output);
                return manifest;
            }
            catch
            {
                if (File.Exists(stage))
                    File.Delete(stage);
                throw;
            }
        }

        private static Dictionary<string, string> GetPayloadSourceFiles(string bundleDirectory)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                {
                    OfflineUpdateLayout.GetPackageManifestPath(),
                    Path.Combine(bundleDirectory, RegulationPackageLayout.ManifestFileName)
                }
            };
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                result.Add(
                    OfflineUpdateLayout.GetModulePath(kind),
                    Path.Combine(
                        bundleDirectory,
                        RegulationPackageLayout.ModulesDirectoryName,
                        RegulationPackageLayout.GetModuleFileName(kind)));
            }
            foreach (string path in result.Values)
            {
                if (!File.Exists(path) || (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidDataException("Payload source khong hop le: " + path + ".");
            }
            return result;
        }

        private static void WriteEntry(ZipArchive zip, string path, byte[] payload)
        {
            ZipArchiveEntry entry = zip.CreateEntry(path, CompressionLevel.Optimal);
            entry.LastWriteTime = DeterministicZipTimestamp;
            using (Stream output = entry.Open())
                output.Write(payload, 0, payload.Length);
        }

        private static void WriteFileEntry(ZipArchive zip, string path, string sourcePath)
        {
            ZipArchiveEntry entry = zip.CreateEntry(path, CompressionLevel.Optimal);
            entry.LastWriteTime = DeterministicZipTimestamp;
            using (Stream input = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (Stream output = entry.Open())
                input.CopyTo(output);
        }

        private static RSACryptoServiceProvider OpenContainer(string containerName, bool create)
        {
            var parameters = new CspParameters(RsaProviderType)
            {
                KeyContainerName = containerName,
                Flags = create
                    ? CspProviderFlags.UseNonExportableKey
                    : CspProviderFlags.UseExistingKey
            };
            var rsa = new RSACryptoServiceProvider(3072, parameters)
            {
                PersistKeyInCsp = true
            };
            return rsa;
        }

        private static void EnsureSamePublicKey(RSAParameters actual, RSAParameters expected)
        {
            if (!actual.Modulus.SequenceEqual(expected.Modulus) ||
                !actual.Exponent.SequenceEqual(expected.Exponent))
            {
                throw new CryptographicException(
                    "Public key file khong khop private key trong Windows key container.");
            }
        }

        private static void ValidateContainerName(string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length < 3 || normalized.Length > 128 ||
                normalized.Any(character => !char.IsLetterOrDigit(character) && character != '-' && character != '_'))
            {
                throw new ArgumentException("Ten Windows key container khong hop le.", nameof(value));
            }
        }

        private static string PrepareNewOutputFile(string path)
        {
            string output = Path.GetFullPath(path ?? string.Empty);
            if (File.Exists(output) || Directory.Exists(output))
                throw new IOException("File dich da ton tai: " + output + ".");
            string parent = Path.GetDirectoryName(output);
            if (string.IsNullOrWhiteSpace(parent))
                throw new ArgumentException("Thu muc dich khong hop le.", nameof(path));
            Directory.CreateDirectory(parent);
            return output;
        }

        private static string ComputeFileChecksum(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (SHA256 sha256 = SHA256.Create())
            {
                return string.Concat(sha256.ComputeHash(stream)
                    .Select(value => value.ToString("X2", CultureInfo.InvariantCulture)));
            }
        }
    }
}
