using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelAddIn1.Core
{
    public sealed class RegulationPackageActivation
    {
        public RegulationPackageActivation(string packageId, string dataVersion, string packageChecksum)
        {
            PackageId = packageId;
            DataVersion = dataVersion;
            PackageChecksum = packageChecksum;
        }

        public string PackageId { get; }
        public string DataVersion { get; }
        public string PackageChecksum { get; }
    }

    public sealed class RegulationPackageActivationStore
    {
        public const string StateFileName = "activation.ttbstate";
        private const string Header = "TTBMVN_PACKAGE_ACTIVATION";
        private readonly string storeRoot;

        public RegulationPackageActivationStore(string storeRoot)
        {
            if (string.IsNullOrWhiteSpace(storeRoot))
                throw new ArgumentException("Package store root la bat buoc.", nameof(storeRoot));
            this.storeRoot = Path.GetFullPath(storeRoot);
        }

        public string StatePath => Path.Combine(storeRoot, StateFileName);

        public IReadOnlyList<RegulationPackageActivation> Load()
        {
            if (!File.Exists(StatePath))
                return new RegulationPackageActivation[0];
            if ((File.GetAttributes(StatePath) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException("Activation state khong duoc la reparse point.");
            string payload = File.ReadAllText(StatePath, new UTF8Encoding(false, true));
            return Deserialize(payload);
        }

        public IReadOnlyList<RegulationPackage> SelectPreferred(
            IEnumerable<RegulationPackage> installedPackages)
        {
            List<RegulationPackage> installed = (installedPackages ?? Enumerable.Empty<RegulationPackage>())
                .Where(package => package != null)
                .ToList();
            Dictionary<string, RegulationPackageActivation> configured = Load()
                .ToDictionary(item => item.PackageId, StringComparer.Ordinal);
            var result = new List<RegulationPackage>();
            foreach (IGrouping<string, RegulationPackage> group in installed
                .GroupBy(package => package.PackageId, StringComparer.Ordinal))
            {
                RegulationPackage selected = null;
                RegulationPackageActivation activation;
                if (configured.TryGetValue(group.Key, out activation))
                {
                    selected = group.FirstOrDefault(package =>
                        string.Equals(package.DataVersion, activation.DataVersion, StringComparison.Ordinal) &&
                        string.Equals(
                            package.PackageChecksum,
                            activation.PackageChecksum,
                            StringComparison.OrdinalIgnoreCase));
                }
                if (selected == null)
                {
                    selected = group.OrderByDescending(
                        package => package.DataVersion,
                        Comparer<string>.Create(RegulationDataVersionComparer.Compare)).First();
                }
                result.Add(selected);
            }
            return new ReadOnlyCollection<RegulationPackage>(result
                .OrderBy(package => package.PackageId, StringComparer.Ordinal)
                .ThenBy(package => package.DataVersion, StringComparer.Ordinal)
                .ToList());
        }

        public void SetPreferred(
            RegulationPackage package,
            IEnumerable<RegulationPackage> installedPackages)
        {
            if (package == null)
                throw new ArgumentNullException(nameof(package));
            List<RegulationPackage> installed = (installedPackages ?? Enumerable.Empty<RegulationPackage>())
                .Where(item => item != null).ToList();
            if (!installed.Any(item => SameIdentity(item, package)))
                throw new InvalidOperationException("Khong the kich hoat package chua duoc cai.");

            var state = Load().ToDictionary(item => item.PackageId, StringComparer.Ordinal);
            state[package.PackageId] = new RegulationPackageActivation(
                package.PackageId,
                package.DataVersion,
                package.PackageChecksum);
            Save(state.Values);
        }

        public bool IsPreferred(
            RegulationPackage package,
            IEnumerable<RegulationPackage> installedPackages)
        {
            if (package == null)
                return false;
            return SelectPreferred(installedPackages).Any(item => SameIdentity(item, package));
        }

        private void Save(IEnumerable<RegulationPackageActivation> activations)
        {
            Directory.CreateDirectory(storeRoot);
            if ((File.GetAttributes(storeRoot) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException("Package store root khong duoc la reparse point.");
            string payload = Serialize(activations);
            string stage = StatePath + ".stage-" + Guid.NewGuid().ToString("N");
            try
            {
                File.WriteAllText(stage, payload, new UTF8Encoding(false));
                if (File.Exists(StatePath))
                    File.Replace(stage, StatePath, null);
                else
                    File.Move(stage, StatePath);
            }
            finally
            {
                if (File.Exists(stage))
                    File.Delete(stage);
            }
        }

        private static string Serialize(IEnumerable<RegulationPackageActivation> activations)
        {
            List<RegulationPackageActivation> items = (activations ?? Enumerable.Empty<RegulationPackageActivation>())
                .OrderBy(item => item.PackageId, StringComparer.Ordinal)
                .ToList();
            Validate(items);
            var lines = new List<string>
            {
                Header,
                "schema=1",
                "count=" + items.Count.ToString(CultureInfo.InvariantCulture)
            };
            for (int index = 0; index < items.Count; index++)
            {
                string prefix = "item." + index.ToString(CultureInfo.InvariantCulture) + ".";
                lines.Add(prefix + "packageId=" + Encode(items[index].PackageId));
                lines.Add(prefix + "dataVersion=" + items[index].DataVersion);
                lines.Add(prefix + "checksum=" + items[index].PackageChecksum.ToUpperInvariant());
            }
            string unsigned = string.Join("\n", lines) + "\n";
            return unsigned + "stateChecksum=" + RegulationPackageSerializer.ComputeSha256(unsigned) + "\n";
        }

        private static IReadOnlyList<RegulationPackageActivation> Deserialize(string payload)
        {
            if (string.IsNullOrEmpty(payload) || payload.Contains("\r") ||
                !payload.EndsWith("\n", StringComparison.Ordinal))
            {
                throw new InvalidDataException("Activation state khong canonical.");
            }
            string[] lines = payload.Split('\n');
            if (lines.Length < 5 || lines[0] != Header || lines[1] != "schema=1")
                throw new InvalidDataException("Activation state header/schema khong hop le.");
            int count = ParsePropertyInt(lines[2], "count");
            int checksumLineIndex = 3 + count * 3;
            if (count < 0 || count > 256 || lines.Length != checksumLineIndex + 2)
                throw new InvalidDataException("Activation state count khong hop le.");
            string checksum = ReadProperty(lines[checksumLineIndex], "stateChecksum");
            string unsigned = string.Join("\n", lines.Take(checksumLineIndex)) + "\n";
            if (!string.Equals(
                checksum,
                RegulationPackageSerializer.ComputeSha256(unsigned),
                StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Activation state checksum khong khop.");
            }

            var result = new List<RegulationPackageActivation>();
            for (int index = 0; index < count; index++)
            {
                int line = 3 + index * 3;
                string prefix = "item." + index.ToString(CultureInfo.InvariantCulture) + ".";
                result.Add(new RegulationPackageActivation(
                    Decode(ReadProperty(lines[line], prefix + "packageId")),
                    ReadProperty(lines[line + 1], prefix + "dataVersion"),
                    ReadProperty(lines[line + 2], prefix + "checksum").ToUpperInvariant()));
            }
            Validate(result);
            if (!string.Equals(Serialize(result), payload, StringComparison.Ordinal))
                throw new InvalidDataException("Activation state khong o dang canonical.");
            return new ReadOnlyCollection<RegulationPackageActivation>(result);
        }

        private static void Validate(IList<RegulationPackageActivation> items)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (RegulationPackageActivation item in items)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.PackageId) || !ids.Add(item.PackageId) ||
                    !RegulationPackageValidator.IsDataVersion(item.DataVersion) ||
                    !RegulationPackageValidator.IsSha256(item.PackageChecksum))
                {
                    throw new InvalidDataException("Activation state co identity khong hop le.");
                }
            }
        }

        private static bool SameIdentity(RegulationPackage left, RegulationPackage right)
        {
            return string.Equals(left.PackageId, right.PackageId, StringComparison.Ordinal) &&
                   string.Equals(left.DataVersion, right.DataVersion, StringComparison.Ordinal) &&
                   string.Equals(left.PackageChecksum, right.PackageChecksum, StringComparison.OrdinalIgnoreCase);
        }

        private static string ReadProperty(string line, string key)
        {
            string prefix = key + "=";
            if (!line.StartsWith(prefix, StringComparison.Ordinal))
                throw new InvalidDataException("Activation state thieu property " + key + ".");
            return line.Substring(prefix.Length);
        }

        private static int ParsePropertyInt(string line, string key)
        {
            int result;
            if (!int.TryParse(ReadProperty(line, key), NumberStyles.None, CultureInfo.InvariantCulture, out result))
                throw new InvalidDataException("Activation state property " + key + " khong hop le.");
            return result;
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value)
        {
            try
            {
                string result = new UTF8Encoding(false, true).GetString(Convert.FromBase64String(value));
                if (!string.Equals(Encode(result), value, StringComparison.Ordinal))
                    throw new FormatException();
                return result;
            }
            catch (Exception ex) when (ex is FormatException || ex is DecoderFallbackException)
            {
                throw new InvalidDataException("Activation state Base64 khong hop le.", ex);
            }
        }
    }
}
