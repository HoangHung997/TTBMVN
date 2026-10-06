using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace ExcelAddIn1.Core
{
    public static class RegulationPackageLayout
    {
        public const string ManifestFileName = "manifest.ttbmanifest";
        public const string ModulesDirectoryName = "modules";
        public const string InstallReceiptFileName = "install.receipt";

        public static string GetModuleFileName(RegulationModuleKind kind)
        {
            if (!Enum.IsDefined(typeof(RegulationModuleKind), kind))
                throw new ArgumentOutOfRangeException(nameof(kind));
            return kind + ".data";
        }
    }

    public enum RegulationPackageInstallStatus
    {
        Installed,
        AlreadyInstalled
    }

    public enum RegulationPackageInstallPhase
    {
        StageValidated,
        BeforeCommit
    }

    public sealed class RegulationPackageInstallResult
    {
        internal RegulationPackageInstallResult(
            RegulationPackageInstallStatus status,
            RegulationPackage package,
            string installDirectory)
        {
            Status = status;
            Package = package;
            InstallDirectory = installDirectory;
        }

        public RegulationPackageInstallStatus Status { get; }
        public RegulationPackage Package { get; }
        public string InstallDirectory { get; }
    }

    public sealed class RegulationPackageVersionConflictException : InvalidOperationException
    {
        public RegulationPackageVersionConflictException(
            string packageId,
            string dataVersion,
            IEnumerable<string> installedChecksums)
            : base("Package " + packageId + " version " + dataVersion +
                   " da ton tai voi checksum khac.")
        {
            PackageId = packageId;
            DataVersion = dataVersion;
            InstalledChecksums = new ReadOnlyCollection<string>(
                (installedChecksums ?? Enumerable.Empty<string>()).ToList());
        }

        public string PackageId { get; }
        public string DataVersion { get; }
        public IReadOnlyList<string> InstalledChecksums { get; }
    }

    public sealed class RegulationPackageStore
    {
        private const long MaximumManifestBytes = 10L * 1024L * 1024L;
        private const int LockAttempts = 200;
        private const int LockRetryMilliseconds = 50;
        private readonly string rootDirectory;
        private readonly Action<RegulationPackageInstallPhase> phaseCallback;

        public RegulationPackageStore(string rootDirectory)
            : this(rootDirectory, null)
        {
        }

        public RegulationPackageStore(
            string rootDirectory,
            Action<RegulationPackageInstallPhase> phaseCallback)
        {
            if (string.IsNullOrWhiteSpace(rootDirectory))
                throw new ArgumentException("Package store root la bat buoc.", nameof(rootDirectory));
            this.rootDirectory = Path.GetFullPath(rootDirectory);
            this.phaseCallback = phaseCallback;
        }

        public string RootDirectory => rootDirectory;

        public RegulationPackageInstallResult ImportFromDirectory(string sourceDirectory)
        {
            string source = ValidateSourceDirectory(sourceDirectory);
            EnsureStoreDirectories();
            string stagingRoot = GetStagingRoot();
            string stage = Path.Combine(stagingRoot, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stage);

            string versionDirectory = null;
            try
            {
                CopyBundleToStage(source, stage);
                RegulationPackage package = ValidateBundle(stage, allowInstallReceipt: false);
                if (package.Status == RegulationPackageStatus.Draft)
                    throw new InvalidDataException("Khong duoc cai package Draft vao kho chinh.");
                phaseCallback?.Invoke(RegulationPackageInstallPhase.StageValidated);

                using (FileStream storeLock = AcquireStoreLock())
                {
                    versionDirectory = GetVersionDirectory(package.PackageId, package.DataVersion);
                    string destination = Path.Combine(versionDirectory, package.PackageChecksum);
                    if (Directory.Exists(destination))
                    {
                        RegulationPackage installed = ValidateBundle(destination, allowInstallReceipt: true);
                        return new RegulationPackageInstallResult(
                            RegulationPackageInstallStatus.AlreadyInstalled,
                            installed,
                            destination);
                    }

                    string[] installedChecksums = Directory.Exists(versionDirectory)
                        ? Directory.GetDirectories(versionDirectory)
                            .Select(Path.GetFileName)
                            .Where(name => !string.IsNullOrWhiteSpace(name))
                            .ToArray()
                        : new string[0];
                    if (installedChecksums.Length > 0)
                    {
                        throw new RegulationPackageVersionConflictException(
                            package.PackageId,
                            package.DataVersion,
                            installedChecksums);
                    }

                    phaseCallback?.Invoke(RegulationPackageInstallPhase.BeforeCommit);
                    WriteInstallReceipt(stage, package);
                    EnsureSafeDirectory(Path.GetDirectoryName(versionDirectory), "Thu muc package");
                    EnsureSafeDirectory(versionDirectory, "Thu muc version");
                    Directory.Move(stage, destination);
                    stage = null;
                    return new RegulationPackageInstallResult(
                        RegulationPackageInstallStatus.Installed,
                        package,
                        destination);
                }
            }
            finally
            {
                if (!string.IsNullOrEmpty(stage) && Directory.Exists(stage))
                    Directory.Delete(stage, true);
                RemoveDirectoryIfEmpty(versionDirectory);
                RemoveDirectoryIfEmpty(versionDirectory == null
                    ? null
                    : Path.GetDirectoryName(versionDirectory));
            }
        }

        public IReadOnlyList<RegulationPackage> ListInstalled()
        {
            string packagesRoot = GetPackagesRoot();
            if (!Directory.Exists(packagesRoot))
                return new RegulationPackage[0];

            var result = new List<RegulationPackage>();
            foreach (string packageDirectory in Directory.GetDirectories(packagesRoot))
            {
                foreach (string versionDirectory in Directory.GetDirectories(packageDirectory))
                {
                    foreach (string checksumDirectory in Directory.GetDirectories(versionDirectory))
                        result.Add(ValidateBundle(checksumDirectory, allowInstallReceipt: true));
                }
            }
            return new ReadOnlyCollection<RegulationPackage>(result
                .OrderBy(package => package.PackageId, StringComparer.Ordinal)
                .ThenBy(package => package.DataVersion, StringComparer.Ordinal)
                .ThenBy(package => package.PackageChecksum, StringComparer.Ordinal)
                .ToList());
        }

        public RegulationPackage LoadRequired(
            string packageId,
            string dataVersion,
            string packageChecksum)
        {
            string directory = GetInstalledPackageDirectory(
                packageId,
                dataVersion,
                packageChecksum);
            if (!Directory.Exists(directory))
                throw new DirectoryNotFoundException("Khong tim thay package da pin trong kho.");
            return ValidateBundle(directory, allowInstallReceipt: true);
        }

        public RegulationPackageBundle LoadBundleRequired(
            string packageId,
            string dataVersion,
            string packageChecksum)
        {
            string directory = GetInstalledPackageDirectory(
                packageId,
                dataVersion,
                packageChecksum);
            if (!Directory.Exists(directory))
                throw new DirectoryNotFoundException("Khong tim thay package da pin trong kho.");
            RegulationPackage validated = ValidateBundle(directory, allowInstallReceipt: true);
            RegulationPackageBundle bundle = RegulationPackageBundleReader.Read(directory);
            if (!string.Equals(validated.PackageId, bundle.Package.PackageId, StringComparison.Ordinal) ||
                !string.Equals(validated.DataVersion, bundle.Package.DataVersion, StringComparison.Ordinal) ||
                !string.Equals(
                    validated.PackageChecksum,
                    bundle.Package.PackageChecksum,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Bundle da doc khong khop package da validate.");
            }
            return bundle;
        }

        private string ValidateSourceDirectory(string sourceDirectory)
        {
            if (string.IsNullOrWhiteSpace(sourceDirectory))
                throw new ArgumentException("Thu muc package nguon la bat buoc.", nameof(sourceDirectory));
            string fullPath = Path.GetFullPath(sourceDirectory);
            if (!Directory.Exists(fullPath))
                throw new DirectoryNotFoundException("Khong tim thay thu muc package nguon: " + fullPath);
            RejectReparsePoint(fullPath, "Thu muc package nguon");
            return fullPath;
        }

        private void EnsureStoreDirectories()
        {
            EnsureSafeDirectory(rootDirectory, "Package store root");
            EnsureSafeDirectory(GetPackagesRoot(), "Thu muc packages");
            EnsureSafeDirectory(GetStagingRoot(), "Thu muc staging");
        }

        private void CopyBundleToStage(string source, string stage)
        {
            string sourceManifest = Path.Combine(source, RegulationPackageLayout.ManifestFileName);
            ValidateRegularFile(sourceManifest, RegulationPackageLayout.ManifestFileName);
            CopyRegularFile(sourceManifest, Path.Combine(stage, RegulationPackageLayout.ManifestFileName));

            string sourceModules = Path.Combine(source, RegulationPackageLayout.ModulesDirectoryName);
            if (!Directory.Exists(sourceModules))
                throw new DirectoryNotFoundException("Package thieu thu muc modules.");
            RejectReparsePoint(sourceModules, "Thu muc modules");
            string stageModules = Path.Combine(stage, RegulationPackageLayout.ModulesDirectoryName);
            Directory.CreateDirectory(stageModules);
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                string fileName = RegulationPackageLayout.GetModuleFileName(kind);
                string sourceFile = Path.Combine(sourceModules, fileName);
                ValidateRegularFile(sourceFile, fileName);
                CopyRegularFile(sourceFile, Path.Combine(stageModules, fileName));
            }
        }

        private RegulationPackage ValidateBundle(string directory, bool allowInstallReceipt)
        {
            if (!Directory.Exists(directory))
                throw new DirectoryNotFoundException("Khong tim thay thu muc package.");
            RejectReparsePoint(directory, "Thu muc package");
            string manifestPath = Path.Combine(directory, RegulationPackageLayout.ManifestFileName);
            ValidateRegularFile(manifestPath, RegulationPackageLayout.ManifestFileName);
            var manifestInfo = new FileInfo(manifestPath);
            if (manifestInfo.Length > MaximumManifestBytes)
                throw new InvalidDataException("Manifest package vuot qua gioi han 10 MB.");
            RegulationPackage package = RegulationPackageSerializer.Deserialize(
                File.ReadAllText(manifestPath, Encoding.UTF8));

            string modulesDirectory = Path.Combine(directory, RegulationPackageLayout.ModulesDirectoryName);
            if (!Directory.Exists(modulesDirectory))
                throw new DirectoryNotFoundException("Package thieu thu muc modules.");
            RejectReparsePoint(modulesDirectory, "Thu muc modules");

            var expectedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (RegulationPackageModuleManifest module in package.Modules)
            {
                string fileName = RegulationPackageLayout.GetModuleFileName(module.Kind);
                expectedFiles.Add(fileName);
                string modulePath = Path.Combine(modulesDirectory, fileName);
                ValidateRegularFile(modulePath, fileName);
                string actualChecksum = ComputeFileChecksum(modulePath);
                if (!string.Equals(
                    actualChecksum,
                    module.ContentChecksum,
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("Checksum module " + module.Kind + " khong khop.");
                }
            }

            string[] actualFiles = Directory.GetFiles(modulesDirectory)
                .Select(Path.GetFileName)
                .ToArray();
            if (actualFiles.Any(file => !expectedFiles.Contains(file)) ||
                actualFiles.Length != expectedFiles.Count)
            {
                throw new InvalidDataException("Thu muc modules co file thua hoac thieu.");
            }

            if (allowInstallReceipt)
                ValidateRegularFile(
                    Path.Combine(directory, RegulationPackageLayout.InstallReceiptFileName),
                    RegulationPackageLayout.InstallReceiptFileName);
            return package;
        }

        private FileStream AcquireStoreLock()
        {
            string lockPath = Path.Combine(rootDirectory, ".store.lock");
            if (File.Exists(lockPath))
                RejectReparsePoint(lockPath, "Package store lock");
            IOException lastError = null;
            for (int attempt = 0; attempt < LockAttempts; attempt++)
            {
                try
                {
                    return new FileStream(
                        lockPath,
                        FileMode.OpenOrCreate,
                        FileAccess.ReadWrite,
                        FileShare.None);
                }
                catch (IOException ex)
                {
                    lastError = ex;
                    Thread.Sleep(LockRetryMilliseconds);
                }
            }
            throw new IOException("Khong lay duoc khoa package store sau 10 giay.", lastError);
        }

        private static void WriteInstallReceipt(string stage, RegulationPackage package)
        {
            string[] lines =
            {
                "TTBMVN_PACKAGE_INSTALL",
                "packageId=" + package.PackageId,
                "dataVersion=" + package.DataVersion,
                "checksum=" + package.PackageChecksum,
                "installedUtc=" + DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
            };
            File.WriteAllText(
                Path.Combine(stage, RegulationPackageLayout.InstallReceiptFileName),
                string.Join("\n", lines),
                new UTF8Encoding(false));
        }

        private string GetPackagesRoot()
        {
            return Path.Combine(rootDirectory, "packages");
        }

        private string GetStagingRoot()
        {
            return Path.Combine(rootDirectory, ".staging");
        }

        private string GetVersionDirectory(string packageId, string dataVersion)
        {
            ValidatePathSegment(packageId, nameof(packageId));
            ValidatePathSegment(dataVersion, nameof(dataVersion));
            return Path.Combine(GetPackagesRoot(), packageId, dataVersion);
        }

        private string GetInstalledPackageDirectory(
            string packageId,
            string dataVersion,
            string packageChecksum)
        {
            ValidatePathSegment(packageId, nameof(packageId));
            ValidatePathSegment(dataVersion, nameof(dataVersion));
            if (!RegulationPackageValidator.IsSha256(packageChecksum))
                throw new ArgumentException("Package checksum khong hop le.", nameof(packageChecksum));
            return Path.Combine(
                GetVersionDirectory(packageId.Trim(), dataVersion.Trim()),
                packageChecksum.Trim().ToUpperInvariant());
        }

        private static void ValidatePathSegment(string value, string parameterName)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Length == 0 ||
                normalized == "." ||
                normalized == ".." ||
                normalized.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                normalized.Contains(Path.DirectorySeparatorChar.ToString()) ||
                normalized.Contains(Path.AltDirectorySeparatorChar.ToString()))
            {
                throw new ArgumentException("Gia tri khong hop le cho path segment.", parameterName);
            }
        }

        private static void ValidateRegularFile(string path, string displayName)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("Package thieu file " + displayName + ".", path);
            RejectReparsePoint(path, "File " + displayName);
        }

        private static void CopyRegularFile(string source, string destination)
        {
            File.Copy(source, destination);
            File.SetAttributes(destination, FileAttributes.Normal);
        }

        private static void EnsureSafeDirectory(string path, string displayName)
        {
            Directory.CreateDirectory(path);
            RejectReparsePoint(path, displayName);
        }

        private static void RejectReparsePoint(string path, string displayName)
        {
            FileAttributes attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException(displayName + " khong duoc la symbolic link/reparse point.");
        }

        private static string ComputeFileChecksum(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(stream);
                var builder = new StringBuilder(hash.Length * 2);
                foreach (byte value in hash)
                    builder.Append(value.ToString("X2", CultureInfo.InvariantCulture));
                return builder.ToString();
            }
        }

        private static void RemoveDirectoryIfEmpty(string path)
        {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
                return;
            if (!Directory.EnumerateFileSystemEntries(path).Any())
                Directory.Delete(path);
        }
    }
}
