using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using ExcelAddIn1.Funtion;

namespace ExcelAddIn1
{
    public static partial class RandomDaodat
    {
        private const int CurrentSettingsSchemaVersion = 2;

        private static void LoadPersistedSettings()
        {
            try
            {
                string text = Properties.Settings.Default.DaodatSettings;
                if (string.IsNullOrWhiteSpace(text))
                    return;

                ApplyPersistedSettingsText(text);
            }
            catch
            {
                settings3m = DaodatRandomSettings.CreateDefault(false);
                settings5m = DaodatRandomSettings.CreateDefault(true);
                runOptions3m = new DaodatRunOptions { Is5m = false };
                runOptions5m = new DaodatRunOptions { Is5m = true };
            }
        }

        private static void SavePersistedSettings()
        {
            try
            {
                Properties.Settings.Default.DaodatSettings = BuildPersistedSettingsText();
                Properties.Settings.Default.Save();
            }
            catch
            {
                // Persisting user settings must not block the main Excel workflow.
            }
        }

        public static void ExportSettingsToFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Duong dan file xuat cai dat khong hop le.");

            File.WriteAllText(filePath, BuildPersistedSettingsText(includeWorkbookFields: false), Encoding.UTF8);
        }

        public static void ImportSettingsFromFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("Khong tim thay file cai dat.", filePath);

            string text = File.ReadAllText(filePath, Encoding.UTF8);
            BackupCurrentPersistedSettings("import");
            ApplyPersistedSettingsText(text, preserveWorkbookFields: true);
            SavePersistedSettings();
        }

        public static string GetDiagnosticSettingsText()
        {
            return BuildPersistedSettingsText(includeWorkbookFields: true);
        }

        public static string GetSanitizedDiagnosticSettingsText()
        {
            return BuildPersistedSettingsText(includeWorkbookFields: false);
        }

        private static string BuildPersistedSettingsText()
        {
            return BuildPersistedSettingsText(includeWorkbookFields: true);
        }

        private static string BuildPersistedSettingsText(bool includeWorkbookFields)
        {
            return BuildJsonPersistedSettingsText(includeWorkbookFields);
        }

        private static string BuildLegacyPersistedSettingsText(bool includeWorkbookFields)
        {
            var lines = new List<string> { "version=1" };
            WritePersistedRandomSettings(lines, "settings3m", settings3m);
            WritePersistedRandomSettings(lines, "settings5m", settings5m);
            WritePersistedRunOptions(lines, "run3m", runOptions3m, includeWorkbookFields);
            WritePersistedRunOptions(lines, "run5m", runOptions5m, includeWorkbookFields);
            return string.Join(Environment.NewLine, lines);
        }

        private static void ApplyPersistedSettingsText(string text)
        {
            ApplyPersistedSettingsText(text, preserveWorkbookFields: false);
        }

        private static void ApplyPersistedSettingsText(string text, bool preserveWorkbookFields)
        {
            if (string.IsNullOrWhiteSpace(text))
                throw new ArgumentException("File cai dat rong.");

            string trimmed = text.TrimStart();
            if (trimmed.StartsWith("{", StringComparison.Ordinal))
            {
                ApplyJsonPersistedSettingsText(text, preserveWorkbookFields);
                return;
            }

            Dictionary<string, string> values = ParsePersistedText(text);
            if (!values.ContainsKey("version"))
                throw new ArgumentException("File cai dat khong dung dinh dang.");

            DaodatRandomSettings importedSettings3m = ReadPersistedRandomSettings(values, "settings3m", DaodatRandomSettings.CreateDefault(false));
            DaodatRandomSettings importedSettings5m = ReadPersistedRandomSettings(values, "settings5m", DaodatRandomSettings.CreateDefault(true));
            DaodatRunOptions importedRunOptions3m = ReadPersistedRunOptions(values, "run3m", false);
            DaodatRunOptions importedRunOptions5m = ReadPersistedRunOptions(values, "run5m", true);

            if (preserveWorkbookFields)
            {
                PreserveWorkbookFields(importedRunOptions3m, runOptions3m);
                PreserveWorkbookFields(importedRunOptions5m, runOptions5m);
            }

            settings3m = importedSettings3m;
            settings5m = importedSettings5m;
            runOptions3m = importedRunOptions3m;
            runOptions5m = importedRunOptions5m;
        }

        private static string BuildJsonPersistedSettingsText(bool includeWorkbookFields)
        {
            var document = new PersistedSettingsDocument
            {
                SchemaVersion = CurrentSettingsSchemaVersion,
                AppVersion = AppInfo.Version,
                CreatedAt = DateTime.Now.ToString("o", CultureInfo.InvariantCulture),
                IncludeWorkbookFields = includeWorkbookFields,
                Settings3m = ToPersistedRandomSettings(settings3m),
                Settings5m = ToPersistedRandomSettings(settings5m),
                Run3m = ToPersistedRunOptions(runOptions3m, includeWorkbookFields),
                Run5m = ToPersistedRunOptions(runOptions5m, includeWorkbookFields)
            };

            using (var stream = new MemoryStream())
            {
                var serializer = new DataContractJsonSerializer(typeof(PersistedSettingsDocument));
                serializer.WriteObject(stream, document);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        private static void ApplyJsonPersistedSettingsText(string text, bool preserveWorkbookFields)
        {
            PersistedSettingsDocument document;
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(text)))
            {
                var serializer = new DataContractJsonSerializer(typeof(PersistedSettingsDocument));
                document = serializer.ReadObject(stream) as PersistedSettingsDocument;
            }

            if (document == null || document.SchemaVersion <= 0)
                throw new ArgumentException("File cai dat khong dung dinh dang JSON.");

            DaodatRandomSettings importedSettings3m = FromPersistedRandomSettings(document.Settings3m, DaodatRandomSettings.CreateDefault(false));
            DaodatRandomSettings importedSettings5m = FromPersistedRandomSettings(document.Settings5m, DaodatRandomSettings.CreateDefault(true));
            DaodatRunOptions importedRunOptions3m = FromPersistedRunOptions(document.Run3m, false);
            DaodatRunOptions importedRunOptions5m = FromPersistedRunOptions(document.Run5m, true);

            if (preserveWorkbookFields)
            {
                PreserveWorkbookFields(importedRunOptions3m, runOptions3m);
                PreserveWorkbookFields(importedRunOptions5m, runOptions5m);
            }

            settings3m = importedSettings3m;
            settings5m = importedSettings5m;
            runOptions3m = importedRunOptions3m;
            runOptions5m = importedRunOptions5m;
        }

        private static void PreserveWorkbookFields(DaodatRunOptions imported, DaodatRunOptions current)
        {
            if (imported == null || current == null)
                return;

            imported.CreateNewSheet = current.CreateNewSheet;
            imported.LinkBack = current.LinkBack;
            imported.UseOutputStartCell = current.UseOutputStartCell;
            imported.OutputStartAddress = current.OutputStartAddress;
            imported.HasTableData = current.HasTableData;
            imported.TableDataAddress = current.TableDataAddress;
            imported.OutputStartRow = current.OutputStartRow;
            imported.NewSheetName = current.NewSheetName;
            imported.ExistingSheetName = current.ExistingSheetName;
            imported.SourceDataColumns = new Dictionary<string, string>(
                current.SourceDataColumns ?? new Dictionary<string, string>(),
                StringComparer.OrdinalIgnoreCase);
        }

        private static Dictionary<string, string> ParsePersistedText(string text)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string rawLine in text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
            {
                if (string.IsNullOrWhiteSpace(rawLine))
                    continue;

                int equalsIndex = rawLine.IndexOf('=');
                if (equalsIndex <= 0)
                    continue;

                values[rawLine.Substring(0, equalsIndex)] = rawLine.Substring(equalsIndex + 1);
            }

            return values;
        }

        private static void WritePersistedRandomSettings(List<string> lines, string prefix, DaodatRandomSettings settings)
        {
            lines.Add($"{prefix}.D1Min={PersistDouble(settings.D1Min)}");
            lines.Add($"{prefix}.D1Max={PersistDouble(settings.D1Max)}");
            lines.Add($"{prefix}.R1Min={PersistDouble(settings.R1Min)}");
            lines.Add($"{prefix}.R1Max={PersistDouble(settings.R1Max)}");
            lines.Add($"{prefix}.D2Min={PersistDouble(settings.D2Min)}");
            lines.Add($"{prefix}.D2Max={PersistDouble(settings.D2Max)}");
            lines.Add($"{prefix}.R2Min={PersistDouble(settings.R2Min)}");
            lines.Add($"{prefix}.R2Max={PersistDouble(settings.R2Max)}");
            lines.Add($"{prefix}.HMin={PersistDouble(settings.HMin)}");
            lines.Add($"{prefix}.HMax={PersistDouble(settings.HMax)}");
            lines.Add($"{prefix}.MaxAttempts={settings.MaxAttempts.ToString(CultureInfo.InvariantCulture)}");
            lines.Add($"{prefix}.MaxParallelWorkers={settings.MaxParallelWorkers.ToString(CultureInfo.InvariantCulture)}");
        }

        private static DaodatRandomSettings ReadPersistedRandomSettings(
            Dictionary<string, string> values,
            string prefix,
            DaodatRandomSettings fallback)
        {
            var settings = fallback.Clone();
            settings.D1Min = ReadPersistedDouble(values, $"{prefix}.D1Min", settings.D1Min);
            settings.D1Max = ReadPersistedDouble(values, $"{prefix}.D1Max", settings.D1Max);
            settings.R1Min = ReadPersistedDouble(values, $"{prefix}.R1Min", settings.R1Min);
            settings.R1Max = ReadPersistedDouble(values, $"{prefix}.R1Max", settings.R1Max);
            settings.D2Min = ReadPersistedDouble(values, $"{prefix}.D2Min", settings.D2Min);
            settings.D2Max = ReadPersistedDouble(values, $"{prefix}.D2Max", settings.D2Max);
            settings.R2Min = ReadPersistedDouble(values, $"{prefix}.R2Min", settings.R2Min);
            settings.R2Max = ReadPersistedDouble(values, $"{prefix}.R2Max", settings.R2Max);
            settings.HMin = ReadPersistedDouble(values, $"{prefix}.HMin", settings.HMin);
            settings.HMax = ReadPersistedDouble(values, $"{prefix}.HMax", settings.HMax);
            settings.MaxAttempts = ReadPersistedInt(values, $"{prefix}.MaxAttempts", settings.MaxAttempts);
            settings.MaxParallelWorkers = ReadPersistedInt(values, $"{prefix}.MaxParallelWorkers", settings.MaxParallelWorkers);
            settings.Validate();
            return settings;
        }

        private static void WritePersistedRunOptions(List<string> lines, string prefix, DaodatRunOptions options, bool includeWorkbookFields)
        {
            options = options ?? new DaodatRunOptions();
            lines.Add($"{prefix}.Is5m={options.Is5m}");
            lines.Add($"{prefix}.InsertRows={options.InsertRows}");
            lines.Add($"{prefix}.ProjectIncludes5m={options.ProjectIncludes5m}");
            lines.Add($"{prefix}.UseDefaultOutputRow={options.UseDefaultOutputRow}");
            lines.Add($"{prefix}.ColumnMappings={EncodePersistedString(SerializeColumnMappings(options.ColumnMappings))}");

            if (!includeWorkbookFields)
                return;

            lines.Add($"{prefix}.CreateNewSheet={options.CreateNewSheet}");
            lines.Add($"{prefix}.LinkBack={options.LinkBack}");
            lines.Add($"{prefix}.UseOutputStartCell={options.UseOutputStartCell}");
            lines.Add($"{prefix}.OutputStartAddress={EncodePersistedString(options.OutputStartAddress)}");
            lines.Add($"{prefix}.HasTableData={options.HasTableData}");
            lines.Add($"{prefix}.TableDataAddress={EncodePersistedString(options.TableDataAddress)}");
            lines.Add($"{prefix}.OutputStartRow={options.OutputStartRow.ToString(CultureInfo.InvariantCulture)}");
            lines.Add($"{prefix}.NewSheetName={EncodePersistedString(options.NewSheetName)}");
            lines.Add($"{prefix}.ExistingSheetName={EncodePersistedString(options.ExistingSheetName)}");
            lines.Add($"{prefix}.SourceDataColumns={EncodePersistedString(SerializeSourceColumns(options.SourceDataColumns))}");
        }

        private static DaodatRunOptions ReadPersistedRunOptions(Dictionary<string, string> values, string prefix, bool is5m)
        {
            return new DaodatRunOptions
            {
                Is5m = is5m,
                InsertRows = ReadPersistedBool(values, $"{prefix}.InsertRows", false),
                CreateNewSheet = ReadPersistedBool(values, $"{prefix}.CreateNewSheet", false),
                LinkBack = ReadPersistedBool(values, $"{prefix}.LinkBack", false),
                UseOutputStartCell = ReadPersistedBool(values, $"{prefix}.UseOutputStartCell", false),
                OutputStartAddress = DecodePersistedString(ReadPersistedString(values, $"{prefix}.OutputStartAddress")),
                HasTableData = ReadPersistedBool(values, $"{prefix}.HasTableData", false),
                ProjectIncludes5m = ReadPersistedBool(values, $"{prefix}.ProjectIncludes5m", false),
                UseDefaultOutputRow = ReadPersistedBool(values, $"{prefix}.UseDefaultOutputRow", true),
                TableDataAddress = DecodePersistedString(ReadPersistedString(values, $"{prefix}.TableDataAddress")),
                OutputStartRow = ReadPersistedInt(values, $"{prefix}.OutputStartRow", 0),
                NewSheetName = DecodePersistedString(ReadPersistedString(values, $"{prefix}.NewSheetName")),
                ExistingSheetName = DecodePersistedString(ReadPersistedString(values, $"{prefix}.ExistingSheetName")),
                SourceDataColumns = DeserializeSourceColumns(DecodePersistedString(ReadPersistedString(values, $"{prefix}.SourceDataColumns"))),
                ColumnMappings = DeserializeColumnMappings(DecodePersistedString(ReadPersistedString(values, $"{prefix}.ColumnMappings")))
            };
        }

        private static PersistedRandomSettings ToPersistedRandomSettings(DaodatRandomSettings settings)
        {
            settings = settings ?? DaodatRandomSettings.CreateDefault(false);
            return new PersistedRandomSettings
            {
                D1Min = settings.D1Min,
                D1Max = settings.D1Max,
                R1Min = settings.R1Min,
                R1Max = settings.R1Max,
                D2Min = settings.D2Min,
                D2Max = settings.D2Max,
                R2Min = settings.R2Min,
                R2Max = settings.R2Max,
                HMin = settings.HMin,
                HMax = settings.HMax,
                MaxAttempts = settings.MaxAttempts,
                MaxParallelWorkers = settings.MaxParallelWorkers
            };
        }

        private static DaodatRandomSettings FromPersistedRandomSettings(PersistedRandomSettings persisted, DaodatRandomSettings fallback)
        {
            DaodatRandomSettings settings = fallback.Clone();
            if (persisted != null)
            {
                settings.D1Min = persisted.D1Min;
                settings.D1Max = persisted.D1Max;
                settings.R1Min = persisted.R1Min;
                settings.R1Max = persisted.R1Max;
                settings.D2Min = persisted.D2Min;
                settings.D2Max = persisted.D2Max;
                settings.R2Min = persisted.R2Min;
                settings.R2Max = persisted.R2Max;
                settings.HMin = persisted.HMin;
                settings.HMax = persisted.HMax;
                settings.MaxAttempts = persisted.MaxAttempts;
                settings.MaxParallelWorkers = persisted.MaxParallelWorkers;
            }

            settings.Validate();
            return settings;
        }

        private static PersistedRunOptions ToPersistedRunOptions(DaodatRunOptions options, bool includeWorkbookFields)
        {
            options = options ?? new DaodatRunOptions();
            var persisted = new PersistedRunOptions
            {
                Is5m = options.Is5m,
                InsertRows = options.InsertRows,
                ProjectIncludes5m = options.ProjectIncludes5m,
                UseDefaultOutputRow = options.UseDefaultOutputRow,
                ColumnMappings = ToPersistedColumnMappings(options.ColumnMappings)
            };

            if (!includeWorkbookFields)
                return persisted;

            persisted.CreateNewSheet = options.CreateNewSheet;
            persisted.LinkBack = options.LinkBack;
            persisted.UseOutputStartCell = options.UseOutputStartCell;
            persisted.OutputStartAddress = options.OutputStartAddress;
            persisted.HasTableData = options.HasTableData;
            persisted.TableDataAddress = options.TableDataAddress;
            persisted.OutputStartRow = options.OutputStartRow;
            persisted.NewSheetName = options.NewSheetName;
            persisted.ExistingSheetName = options.ExistingSheetName;
            persisted.SourceDataColumns = ToPersistedSourceColumns(options.SourceDataColumns);
            return persisted;
        }

        private static DaodatRunOptions FromPersistedRunOptions(PersistedRunOptions persisted, bool is5m)
        {
            if (persisted == null)
                return new DaodatRunOptions { Is5m = is5m };

            return new DaodatRunOptions
            {
                Is5m = is5m,
                InsertRows = persisted.InsertRows,
                CreateNewSheet = persisted.CreateNewSheet,
                LinkBack = persisted.LinkBack,
                UseOutputStartCell = persisted.UseOutputStartCell,
                OutputStartAddress = persisted.OutputStartAddress,
                HasTableData = persisted.HasTableData,
                ProjectIncludes5m = persisted.ProjectIncludes5m,
                UseDefaultOutputRow = persisted.UseDefaultOutputRow,
                TableDataAddress = persisted.TableDataAddress,
                OutputStartRow = persisted.OutputStartRow,
                NewSheetName = persisted.NewSheetName,
                ExistingSheetName = persisted.ExistingSheetName,
                SourceDataColumns = FromPersistedSourceColumns(persisted.SourceDataColumns),
                ColumnMappings = FromPersistedColumnMappings(persisted.ColumnMappings)
            };
        }

        private static List<PersistedKeyValue> ToPersistedSourceColumns(Dictionary<string, string> columns)
        {
            return (columns ?? new Dictionary<string, string>())
                .Select(x => new PersistedKeyValue { Key = x.Key, Value = x.Value })
                .ToList();
        }

        private static Dictionary<string, string> FromPersistedSourceColumns(List<PersistedKeyValue> columns)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (PersistedKeyValue item in columns ?? new List<PersistedKeyValue>())
            {
                if (!string.IsNullOrWhiteSpace(item.Key))
                    result[item.Key] = item.Value ?? string.Empty;
            }

            return result;
        }

        private static List<PersistedColumnMapping> ToPersistedColumnMappings(List<DaodatColumnMapping> mappings)
        {
            return (mappings ?? new List<DaodatColumnMapping>())
                .Select(x => new PersistedColumnMapping
                {
                    Key = x.Key,
                    SourceColumn = x.SourceColumn,
                    OutputColumn = x.OutputColumn,
                    FormatLocal = x.FormatLocal
                })
                .ToList();
        }

        private static List<DaodatColumnMapping> FromPersistedColumnMappings(List<PersistedColumnMapping> mappings)
        {
            return (mappings ?? new List<PersistedColumnMapping>())
                .Select(x => new DaodatColumnMapping
                {
                    Key = x.Key,
                    SourceColumn = x.SourceColumn,
                    OutputColumn = x.OutputColumn,
                    FormatLocal = x.FormatLocal
                })
                .ToList();
        }

        private static string SerializeSourceColumns(Dictionary<string, string> columns)
        {
            if (columns == null || columns.Count == 0)
                return string.Empty;

            return string.Join("\n", columns.Select(x => $"{EncodePersistedString(x.Key)}\t{EncodePersistedString(x.Value)}"));
        }

        private static Dictionary<string, string> DeserializeSourceColumns(string text)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(text))
                return result;

            foreach (string line in text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] parts = line.Split('\t');
                if (parts.Length >= 2)
                    result[DecodePersistedString(parts[0])] = DecodePersistedString(parts[1]);
            }

            return result;
        }

        private static string SerializeColumnMappings(List<DaodatColumnMapping> mappings)
        {
            if (mappings == null || mappings.Count == 0)
                return string.Empty;

            return string.Join("\n", mappings.Select(x =>
                $"{EncodePersistedString(x.Key)}\t{EncodePersistedString(x.SourceColumn)}\t{EncodePersistedString(x.OutputColumn)}\t{EncodePersistedString(x.FormatLocal)}"));
        }

        private static List<DaodatColumnMapping> DeserializeColumnMappings(string text)
        {
            var result = new List<DaodatColumnMapping>();
            if (string.IsNullOrWhiteSpace(text))
                return result;

            foreach (string line in text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string[] parts = line.Split('\t');
                result.Add(new DaodatColumnMapping
                {
                    Key = parts.Length > 0 ? DecodePersistedString(parts[0]) : string.Empty,
                    SourceColumn = parts.Length > 1 ? DecodePersistedString(parts[1]) : string.Empty,
                    OutputColumn = parts.Length > 2 ? DecodePersistedString(parts[2]) : string.Empty,
                    FormatLocal = parts.Length > 3 ? DecodePersistedString(parts[3]) : string.Empty
                });
            }

            return result;
        }

        private static string PersistDouble(double value)
        {
            return value.ToString("R", CultureInfo.InvariantCulture);
        }

        private static double ReadPersistedDouble(Dictionary<string, string> values, string key, double fallback)
        {
            return values.TryGetValue(key, out string text)
                && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? value
                : fallback;
        }

        private static int ReadPersistedInt(Dictionary<string, string> values, string key, int fallback)
        {
            return values.TryGetValue(key, out string text)
                && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)
                ? value
                : fallback;
        }

        private static bool ReadPersistedBool(Dictionary<string, string> values, string key, bool fallback)
        {
            return values.TryGetValue(key, out string text)
                && bool.TryParse(text, out bool value)
                ? value
                : fallback;
        }

        private static string ReadPersistedString(Dictionary<string, string> values, string key)
        {
            return values.TryGetValue(key, out string value) ? value : string.Empty;
        }

        private static string EncodePersistedString(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string DecodePersistedString(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value));
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void BackupCurrentPersistedSettings(string reason)
        {
            try
            {
                string text = Properties.Settings.Default.DaodatSettings;
                if (string.IsNullOrWhiteSpace(text))
                    return;

                Directory.CreateDirectory(AppPaths.SettingsBackupDirectory);
                string safeReason = string.IsNullOrWhiteSpace(reason) ? "backup" : reason.Trim();
                string fileName = "daodat_" + safeReason + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".bak";
                File.WriteAllText(Path.Combine(AppPaths.SettingsBackupDirectory, fileName), text, Encoding.UTF8);
            }
            catch
            {
                // Backup is best-effort; importing settings should still be able to continue.
            }
        }

        [DataContract]
        private sealed class PersistedSettingsDocument
        {
            [DataMember(Order = 1)]
            public int SchemaVersion { get; set; }

            [DataMember(Order = 2)]
            public string AppVersion { get; set; }

            [DataMember(Order = 3)]
            public string CreatedAt { get; set; }

            [DataMember(Order = 4)]
            public bool IncludeWorkbookFields { get; set; }

            [DataMember(Order = 5)]
            public PersistedRandomSettings Settings3m { get; set; }

            [DataMember(Order = 6)]
            public PersistedRandomSettings Settings5m { get; set; }

            [DataMember(Order = 7)]
            public PersistedRunOptions Run3m { get; set; }

            [DataMember(Order = 8)]
            public PersistedRunOptions Run5m { get; set; }
        }

        [DataContract]
        private sealed class PersistedRandomSettings
        {
            [DataMember(Order = 1)] public double D1Min { get; set; }
            [DataMember(Order = 2)] public double D1Max { get; set; }
            [DataMember(Order = 3)] public double R1Min { get; set; }
            [DataMember(Order = 4)] public double R1Max { get; set; }
            [DataMember(Order = 5)] public double D2Min { get; set; }
            [DataMember(Order = 6)] public double D2Max { get; set; }
            [DataMember(Order = 7)] public double R2Min { get; set; }
            [DataMember(Order = 8)] public double R2Max { get; set; }
            [DataMember(Order = 9)] public double HMin { get; set; }
            [DataMember(Order = 10)] public double HMax { get; set; }
            [DataMember(Order = 11)] public int MaxAttempts { get; set; }
            [DataMember(Order = 12)] public int MaxParallelWorkers { get; set; }
        }

        [DataContract]
        private sealed class PersistedRunOptions
        {
            [DataMember(Order = 1)] public bool Is5m { get; set; }
            [DataMember(Order = 2)] public bool InsertRows { get; set; }
            [DataMember(Order = 3)] public bool CreateNewSheet { get; set; }
            [DataMember(Order = 4)] public bool LinkBack { get; set; }
            [DataMember(Order = 5)] public bool UseOutputStartCell { get; set; }
            [DataMember(Order = 6)] public string OutputStartAddress { get; set; }
            [DataMember(Order = 7)] public bool HasTableData { get; set; }
            [DataMember(Order = 8)] public bool ProjectIncludes5m { get; set; }
            [DataMember(Order = 9)] public bool UseDefaultOutputRow { get; set; }
            [DataMember(Order = 10)] public string TableDataAddress { get; set; }
            [DataMember(Order = 11)] public int OutputStartRow { get; set; }
            [DataMember(Order = 12)] public string NewSheetName { get; set; }
            [DataMember(Order = 13)] public string ExistingSheetName { get; set; }
            [DataMember(Order = 14)] public List<PersistedKeyValue> SourceDataColumns { get; set; }
            [DataMember(Order = 15)] public List<PersistedColumnMapping> ColumnMappings { get; set; }
        }

        [DataContract]
        private sealed class PersistedKeyValue
        {
            [DataMember(Order = 1)] public string Key { get; set; }
            [DataMember(Order = 2)] public string Value { get; set; }
        }

        [DataContract]
        private sealed class PersistedColumnMapping
        {
            [DataMember(Order = 1)] public string Key { get; set; }
            [DataMember(Order = 2)] public string SourceColumn { get; set; }
            [DataMember(Order = 3)] public string OutputColumn { get; set; }
            [DataMember(Order = 4)] public string FormatLocal { get; set; }
        }
    }
}
