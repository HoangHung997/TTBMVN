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
    public sealed class RegulationPackageSourceDefinition
    {
        internal RegulationPackageSourceDefinition(
            string packageId,
            string dataVersion,
            DateTime effectiveFrom,
            DateTime? effectiveTo,
            RegulationPackageStatus status,
            string transitionNote,
            IEnumerable<RegulationPackageSourceDocument> sources,
            IDictionary<RegulationModuleKind, IReadOnlyList<RegulationDataRecord>> records)
        {
            PackageId = packageId;
            DataVersion = dataVersion;
            EffectiveFrom = effectiveFrom;
            EffectiveTo = effectiveTo;
            Status = status;
            TransitionNote = transitionNote;
            Sources = new ReadOnlyCollection<RegulationPackageSourceDocument>(sources.ToList());
            Records = new ReadOnlyDictionary<RegulationModuleKind, IReadOnlyList<RegulationDataRecord>>(
                new Dictionary<RegulationModuleKind, IReadOnlyList<RegulationDataRecord>>(records));
        }

        public string PackageId { get; }
        public string DataVersion { get; }
        public DateTime EffectiveFrom { get; }
        public DateTime? EffectiveTo { get; }
        public RegulationPackageStatus Status { get; }
        public string TransitionNote { get; }
        public IReadOnlyList<RegulationPackageSourceDocument> Sources { get; }
        public IReadOnlyDictionary<RegulationModuleKind, IReadOnlyList<RegulationDataRecord>> Records { get; }
    }

    public static class RegulationPackageSourceReader
    {
        public const string PropertiesFileName = "package.properties";
        public const string SourcesFileName = "sources.tsv";
        public const string SourceModulesDirectoryName = "modules";

        private static readonly string[] SourceColumns =
        {
            "documentId", "title", "publisher", "issuedDate", "effectiveFrom",
            "effectiveTo", "officialUri", "contentChecksum"
        };

        private static readonly string[] RecordColumns =
        {
            "key", "recordType", "unit", "title", "data", "sourceDocumentId",
            "pageFrom", "pageTo", "section", "verification"
        };

        public static RegulationPackageSourceDefinition Read(string sourceDirectory)
        {
            string root = RequireDirectory(sourceDirectory);
            Dictionary<string, string> properties = ReadProperties(
                Path.Combine(root, PropertiesFileName));
            RequireExactKeys(
                properties,
                "packageId", "dataVersion", "effectiveFrom", "effectiveTo", "status", "transitionNote");

            var sources = new List<RegulationPackageSourceDocument>();
            foreach (string[] row in ReadTsv(Path.Combine(root, SourcesFileName), SourceColumns))
            {
                sources.Add(new RegulationPackageSourceDocument(
                    row[0], row[1], row[2],
                    ParseDate(row[3], "issuedDate"),
                    ParseDate(row[4], "effectiveFrom"),
                    ParseNullableDate(row[5], "effectiveTo"),
                    row[6], row[7]));
            }

            string moduleRoot = RequireDirectory(Path.Combine(root, SourceModulesDirectoryName));
            var records = new Dictionary<RegulationModuleKind, IReadOnlyList<RegulationDataRecord>>();
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                string path = Path.Combine(moduleRoot, kind + ".tsv");
                var moduleRecords = new List<RegulationDataRecord>();
                foreach (string[] row in ReadTsv(path, RecordColumns))
                {
                    RegulationDataVerification verification;
                    if (!Enum.TryParse(row[9], false, out verification))
                        throw new InvalidDataException("Verification khong hop le trong " + path + ".");
                    moduleRecords.Add(new RegulationDataRecord(
                        row[0], row[1], row[2], row[3], row[4],
                        new RegulationSourceLocator(
                            row[5],
                            ParsePositiveInt(row[6], "pageFrom", path),
                            ParsePositiveInt(row[7], "pageTo", path),
                            row[8]),
                        verification));
                }
                records.Add(kind, new ReadOnlyCollection<RegulationDataRecord>(moduleRecords));
            }

            return new RegulationPackageSourceDefinition(
                Get(properties, "packageId"),
                Get(properties, "dataVersion"),
                ParseDate(Get(properties, "effectiveFrom"), "effectiveFrom"),
                ParseNullableDate(Get(properties, "effectiveTo"), "effectiveTo"),
                ParseStatus(Get(properties, "status")),
                Get(properties, "transitionNote"),
                sources,
                records);
        }

        private static Dictionary<string, string> ReadProperties(string path)
        {
            RequireRegularFile(path);
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            int lineNumber = 0;
            foreach (string rawLine in File.ReadAllLines(path, StrictUtf8()))
            {
                lineNumber++;
                string line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
                    continue;
                int separator = line.IndexOf('=');
                if (separator <= 0)
                    throw new InvalidDataException("Dong properties khong hop le: " + path + ":" + lineNumber + ".");
                string key = line.Substring(0, separator).Trim();
                string value = line.Substring(separator + 1).Trim();
                if (result.ContainsKey(key))
                    throw new InvalidDataException("Trung property " + key + " trong " + path + ".");
                result.Add(key, value);
            }
            return result;
        }

        private static IEnumerable<string[]> ReadTsv(string path, string[] expectedColumns)
        {
            RequireRegularFile(path);
            string[] lines = File.ReadAllLines(path, StrictUtf8());
            if (lines.Length < 2)
                throw new InvalidDataException("TSV khong co record: " + path + ".");
            string expectedHeader = string.Join("\t", expectedColumns);
            if (!string.Equals(lines[0].TrimStart('\uFEFF'), expectedHeader, StringComparison.Ordinal))
                throw new InvalidDataException("Header TSV khong hop le: " + path + ".");

            for (int index = 1; index < lines.Length; index++)
            {
                if (string.IsNullOrWhiteSpace(lines[index]))
                    throw new InvalidDataException("TSV co dong trong: " + path + ":" + (index + 1) + ".");
                string[] fields = lines[index].Split(new[] { '\t' }, StringSplitOptions.None);
                if (fields.Length != expectedColumns.Length)
                    throw new InvalidDataException("So cot TSV khong hop le: " + path + ":" + (index + 1) + ".");
                for (int fieldIndex = 0; fieldIndex < fields.Length; fieldIndex++)
                    fields[fieldIndex] = fields[fieldIndex].Trim();
                yield return fields;
            }
        }

        private static void RequireExactKeys(IDictionary<string, string> values, params string[] expected)
        {
            var keys = new HashSet<string>(values.Keys, StringComparer.Ordinal);
            foreach (string key in expected)
            {
                if (!keys.Remove(key))
                    throw new InvalidDataException("Thieu property " + key + ".");
            }
            if (keys.Count > 0)
                throw new InvalidDataException("Property khong duoc ho tro: " + string.Join(", ", keys) + ".");
        }

        private static string Get(IDictionary<string, string> values, string key)
        {
            string value;
            if (!values.TryGetValue(key, out value))
                throw new InvalidDataException("Thieu property " + key + ".");
            return value;
        }

        private static RegulationPackageStatus ParseStatus(string value)
        {
            RegulationPackageStatus status;
            if (!Enum.TryParse(value, false, out status) || !Enum.IsDefined(typeof(RegulationPackageStatus), status))
                throw new InvalidDataException("Package status khong hop le.");
            return status;
        }

        private static DateTime ParseDate(string value, string fieldName)
        {
            DateTime result;
            if (!DateTime.TryParseExact(
                value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
            {
                throw new InvalidDataException(fieldName + " phai dung dang yyyy-MM-dd.");
            }
            return result;
        }

        private static DateTime? ParseNullableDate(string value, string fieldName)
        {
            return string.IsNullOrWhiteSpace(value) ? (DateTime?)null : ParseDate(value, fieldName);
        }

        private static int ParsePositiveInt(string value, string fieldName, string path)
        {
            int result;
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out result) || result < 1)
                throw new InvalidDataException(fieldName + " khong hop le trong " + path + ".");
            return result;
        }

        private static string RequireDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Duong dan thu muc trong.", nameof(path));
            string fullPath = Path.GetFullPath(path);
            if (!Directory.Exists(fullPath))
                throw new DirectoryNotFoundException(fullPath);
            return fullPath;
        }

        private static void RequireRegularFile(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("Khong tim thay file nguon.", path);
            FileAttributes attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.ReparsePoint) != 0 ||
                (attributes & FileAttributes.Directory) != 0)
            {
                throw new InvalidDataException("File nguon khong phai regular file: " + path + ".");
            }
        }

        private static UTF8Encoding StrictUtf8()
        {
            return new UTF8Encoding(false, true);
        }
    }

    public sealed class RegulationPackageBundle
    {
        internal RegulationPackageBundle(
            RegulationPackage package,
            IDictionary<RegulationModuleKind, RegulationDataModule> modules,
            string directory)
        {
            Package = package;
            Modules = new ReadOnlyDictionary<RegulationModuleKind, RegulationDataModule>(
                new Dictionary<RegulationModuleKind, RegulationDataModule>(modules));
            Directory = directory;
        }

        public RegulationPackage Package { get; }
        public IReadOnlyDictionary<RegulationModuleKind, RegulationDataModule> Modules { get; }
        public string Directory { get; }
    }

    public static class RegulationPackageBundleReader
    {
        public static RegulationPackageBundle Read(string bundleDirectory)
        {
            if (string.IsNullOrWhiteSpace(bundleDirectory))
                throw new ArgumentException("Duong dan bundle trong.", nameof(bundleDirectory));
            string root = Path.GetFullPath(bundleDirectory);
            if (!Directory.Exists(root))
                throw new DirectoryNotFoundException(root);

            RegulationPackage package = RegulationPackageSerializer.Deserialize(
                File.ReadAllText(
                    Path.Combine(root, RegulationPackageLayout.ManifestFileName),
                    new UTF8Encoding(false, true)));
            string modulesDirectory = Path.Combine(root, RegulationPackageLayout.ModulesDirectoryName);
            var modules = new Dictionary<RegulationModuleKind, RegulationDataModule>();
            foreach (RegulationPackageModuleManifest manifest in package.Modules)
            {
                string path = Path.Combine(
                    modulesDirectory,
                    RegulationPackageLayout.GetModuleFileName(manifest.Kind));
                string fileChecksum = ComputeFileChecksum(path);
                if (!string.Equals(fileChecksum, manifest.ContentChecksum, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Checksum file module khong khop manifest: " + manifest.Kind + ".");
                RegulationDataModule module = RegulationDataModuleSerializer.Deserialize(
                    File.ReadAllText(path, new UTF8Encoding(false, true)));
                if (module.Kind != manifest.Kind ||
                    module.SchemaVersion != manifest.SchemaVersion ||
                    !string.Equals(module.DataVersion, manifest.DataVersion, StringComparison.Ordinal) ||
                    module.Records.Count != manifest.RecordCount)
                {
                    throw new InvalidDataException("Module khong khop manifest: " + manifest.Kind + ".");
                }
                RegulationDataValidationResult validation = RegulationDataValidator.ValidateAgainstPackage(
                    module, package, true);
                if (!validation.IsValid)
                    throw new InvalidDataException(string.Join(" ", validation.Errors));
                modules.Add(module.Kind, module);
            }

            string[] actualFiles = Directory.GetFiles(modulesDirectory)
                .Select(Path.GetFileName)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            string[] expectedFiles = package.Modules
                .Select(module => RegulationPackageLayout.GetModuleFileName(module.Kind))
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();
            if (!actualFiles.SequenceEqual(expectedFiles, StringComparer.Ordinal))
                throw new InvalidDataException("Thu muc module co file thieu hoac thua.");

            return new RegulationPackageBundle(package, modules, root);
        }

        private static string ComputeFileChecksum(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (SHA256 sha256 = SHA256.Create())
            {
                return BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", string.Empty);
            }
        }
    }

    public static class RegulationPackageBundleBuilder
    {
        public static RegulationPackageBundle Build(string sourceDirectory, string outputDirectory)
        {
            RegulationPackageSourceDefinition source = RegulationPackageSourceReader.Read(sourceDirectory);
            var modules = new Dictionary<RegulationModuleKind, RegulationDataModule>();
            var modulePayloads = new Dictionary<RegulationModuleKind, string>();
            var manifests = new List<RegulationPackageModuleManifest>();
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                RegulationDataModule module = RegulationDataModule.Create(
                    kind, source.DataVersion, source.Records[kind]);
                modules.Add(kind, module);
                string payload = RegulationDataModuleSerializer.Serialize(module);
                modulePayloads.Add(kind, payload);
                manifests.Add(new RegulationPackageModuleManifest(
                    kind,
                    source.PackageId + "-" + kind,
                    module.SchemaVersion,
                    module.DataVersion,
                    module.Records.Count,
                    RegulationPackageSerializer.ComputeSha256(payload)));
            }

            RegulationPackage package = RegulationPackage.Create(
                source.PackageId,
                source.DataVersion,
                source.EffectiveFrom,
                source.EffectiveTo,
                source.Status,
                source.TransitionNote,
                source.Sources,
                manifests);
            foreach (RegulationDataModule module in modules.Values)
            {
                RegulationDataValidationResult validation = RegulationDataValidator.ValidateAgainstPackage(
                    module, package, true);
                if (!validation.IsValid)
                    throw new InvalidDataException(string.Join(" ", validation.Errors));
            }

            string destination = Path.GetFullPath(outputDirectory);
            if (Directory.Exists(destination))
            {
                RegulationPackageBundle existing = RegulationPackageBundleReader.Read(destination);
                if (string.Equals(
                    existing.Package.PackageChecksum,
                    package.PackageChecksum,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return existing;
                }
                throw new IOException("Bundle dich da ton tai voi checksum khac: " + destination + ".");
            }

            string parent = Path.GetDirectoryName(destination);
            if (string.IsNullOrWhiteSpace(parent))
                throw new ArgumentException("Thu muc dich khong hop le.", nameof(outputDirectory));
            Directory.CreateDirectory(parent);
            string stage = Path.Combine(parent, ".bundle-stage-" + Guid.NewGuid().ToString("N"));
            try
            {
                string moduleDirectory = Path.Combine(stage, RegulationPackageLayout.ModulesDirectoryName);
                Directory.CreateDirectory(moduleDirectory);
                foreach (KeyValuePair<RegulationModuleKind, RegulationDataModule> item in modules)
                {
                    File.WriteAllText(
                        Path.Combine(moduleDirectory, RegulationPackageLayout.GetModuleFileName(item.Key)),
                        modulePayloads[item.Key],
                        new UTF8Encoding(false));
                }
                File.WriteAllText(
                    Path.Combine(stage, RegulationPackageLayout.ManifestFileName),
                    RegulationPackageSerializer.Serialize(package),
                    new UTF8Encoding(false));
                RegulationPackageBundleReader.Read(stage);
                Directory.Move(stage, destination);
                return RegulationPackageBundleReader.Read(destination);
            }
            catch
            {
                if (Directory.Exists(stage))
                    Directory.Delete(stage, true);
                throw;
            }
        }
    }
}
