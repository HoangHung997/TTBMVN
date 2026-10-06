using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAddIn1.Core
{
    public enum PriceProfileInstallStatus
    {
        Installed = 1,
        AlreadyInstalled = 2
    }

    public sealed class PriceProfileInstallResult
    {
        internal PriceProfileInstallResult(
            PriceProfileInstallStatus status,
            PriceProfile profile,
            string directory)
        {
            Status = status;
            Profile = profile;
            Directory = directory;
        }

        public PriceProfileInstallStatus Status { get; }
        public PriceProfile Profile { get; }
        public string Directory { get; }
    }

    public sealed class PriceProfileStore
    {
        private const string ProfileFileName = "profile.ttbprice";
        private readonly string root;

        public PriceProfileStore(string rootDirectory)
        {
            if (string.IsNullOrWhiteSpace(rootDirectory))
                throw new ArgumentException("Thu muc PriceProfile store trong.", nameof(rootDirectory));
            root = Path.GetFullPath(rootDirectory);
        }

        public PriceProfileInstallResult Import(PriceProfile profile)
        {
            PriceProfileValidator.ValidateRequired(profile, verifyChecksum: true);
            Directory.CreateDirectory(root);
            string target = GetVersionDirectory(profile.ProfileId, profile.DataVersion);
            if (Directory.Exists(target))
                return ResolveExisting(target, profile);

            string staging = Path.Combine(root, ".staging-" + Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(staging);
                string file = Path.Combine(staging, ProfileFileName);
                File.WriteAllText(file, PriceProfileSerializer.Serialize(profile), new UTF8Encoding(false));
                PriceProfile verified = LoadFile(file);
                if (!string.Equals(verified.Checksum, profile.Checksum, StringComparison.Ordinal))
                    throw new InvalidDataException("PriceProfile staging checksum khong khop.");

                string parent = Path.GetDirectoryName(target);
                Directory.CreateDirectory(parent);
                try
                {
                    Directory.Move(staging, target);
                    staging = null;
                    return new PriceProfileInstallResult(
                        PriceProfileInstallStatus.Installed,
                        verified,
                        target);
                }
                catch (IOException)
                {
                    if (!Directory.Exists(target))
                        throw;
                    return ResolveExisting(target, profile);
                }
            }
            finally
            {
                if (!string.IsNullOrEmpty(staging) && Directory.Exists(staging))
                    Directory.Delete(staging, true);
            }
        }

        public PriceProfileInstallResult ImportFile(string filePath)
        {
            return Import(LoadFile(filePath));
        }

        public PriceProfile LoadRequired(string profileId, string dataVersion, string checksum)
        {
            string directory = GetVersionDirectory(profileId, dataVersion);
            if (!Directory.Exists(directory))
                throw new DirectoryNotFoundException("Khong co PriceProfile " + profileId + " v" + dataVersion + ".");
            PriceProfile profile = LoadFile(Path.Combine(directory, ProfileFileName));
            if (!string.Equals(profile.Checksum, (checksum ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("PriceProfile checksum khong khop identity yeu cau.");
            return profile;
        }

        public IReadOnlyList<PriceProfile> ListInstalled()
        {
            if (!Directory.Exists(root))
                return new ReadOnlyCollection<PriceProfile>(new List<PriceProfile>());
            var result = new List<PriceProfile>();
            foreach (string idDirectory in Directory.GetDirectories(root))
            {
                if (Path.GetFileName(idDirectory).StartsWith(".staging-", StringComparison.OrdinalIgnoreCase))
                    continue;
                foreach (string versionDirectory in Directory.GetDirectories(idDirectory))
                {
                    string file = Path.Combine(versionDirectory, ProfileFileName);
                    if (File.Exists(file))
                        result.Add(LoadFile(file));
                }
            }
            return new ReadOnlyCollection<PriceProfile>(result
                .OrderBy(item => item.ProfileId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.DataVersion, StringComparer.Ordinal)
                .ToList());
        }

        public static PriceProfile LoadFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Duong dan PriceProfile trong.", nameof(filePath));
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException("Khong tim thay PriceProfile.", fullPath);
            return PriceProfileSerializer.Deserialize(
                File.ReadAllText(fullPath, new UTF8Encoding(false, true)));
        }

        private PriceProfileInstallResult ResolveExisting(string directory, PriceProfile incoming)
        {
            string file = Path.Combine(directory, ProfileFileName);
            if (!File.Exists(file))
                throw new InvalidDataException("Thu muc PriceProfile da ton tai nhung thieu payload.");
            PriceProfile existing = LoadFile(file);
            if (!string.Equals(existing.ProfileId, incoming.ProfileId, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(existing.DataVersion, incoming.DataVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException("Identity PriceProfile trong store khong khop duong dan.");
            }
            if (!string.Equals(existing.Checksum, incoming.Checksum, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "PriceProfile cung ID/version da ton tai voi checksum khac; hay tang version.");
            }
            return new PriceProfileInstallResult(
                PriceProfileInstallStatus.AlreadyInstalled,
                existing,
                directory);
        }

        private string GetVersionDirectory(string profileId, string dataVersion)
        {
            string id = (profileId ?? string.Empty).Trim();
            string version = (dataVersion ?? string.Empty).Trim();
            if (id.Length == 0 || version.Length == 0 ||
                id.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                version.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                id == "." || id == ".." || version == "." || version == "..")
            {
                throw new ArgumentException("Identity PriceProfile khong hop le.");
            }
            string path = Path.GetFullPath(Path.Combine(root, id, version));
            string prefix = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) +
                Path.DirectorySeparatorChar;
            if (!path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Duong dan PriceProfile vuot store root.");
            return path;
        }
    }
}
