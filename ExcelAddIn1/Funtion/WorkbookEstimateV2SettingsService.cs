using ExcelAddIn1.Core;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public enum EstimateV2OutputSlot
    {
        CostSummary = 0,
        EstimateAppendix = 1,
        UnitRateLand = 2,
        UnitRateWater = 3,
        ResourcePrices = 4,
        UnitRateSea = 5
    }

    public sealed class WorkbookEstimateV2SettingsSnapshot
    {
        internal WorkbookEstimateV2SettingsSnapshot(
            EstimateV2Settings settings,
            IEnumerable<WorkbookSheetDescriptor> sheets,
            IReadOnlyDictionary<EstimateV2OutputSlot, string> sheetKeys,
            bool metadataValid,
            int mappingLossCount)
        {
            Settings = settings ?? throw new ArgumentNullException(nameof(settings));
            Sheets = new ReadOnlyCollection<WorkbookSheetDescriptor>(
                (sheets ?? Enumerable.Empty<WorkbookSheetDescriptor>())
                    .Where(item => item != null)
                    .ToList());
            SheetKeys = new ReadOnlyDictionary<EstimateV2OutputSlot, string>(
                sheetKeys == null
                    ? new Dictionary<EstimateV2OutputSlot, string>()
                    : sheetKeys.ToDictionary(
                        item => item.Key,
                        item => item.Value));
            MetadataValid = metadataValid;
            MappingLossCount = Math.Max(0, mappingLossCount);
        }

        public EstimateV2Settings Settings { get; }
        public IReadOnlyList<WorkbookSheetDescriptor> Sheets { get; }
        public IReadOnlyDictionary<EstimateV2OutputSlot, string> SheetKeys { get; }
        public bool MetadataValid { get; }
        public int MappingLossCount { get; }
        public bool MappingLossDetected => MappingLossCount > 0;

        public int ConfiguredSheetCount =>
            SheetKeys.Values.Count(value =>
                !string.IsNullOrWhiteSpace(value));
    }

    public static class WorkbookEstimateV2SettingsService
    {
        private const string Prefix =
            "TTBMVN.EstimateV2.Settings.";
        private const string AutoRestoreNorm =
            Prefix + "AutoRestoreNorm";
        private const string ValidateOnOpen =
            Prefix + "ValidateOnOpen";
        private const string FormulaLinks =
            Prefix + "FormulaLinks";
        private const string AutoSyncRows =
            Prefix + "AutoSyncRows";
        private const string UseCustomXml =
            Prefix + "UseCustomXml";
        private const string HideTechnical =
            Prefix + "HideTechnical";
        private const string WarnMappingLoss =
            Prefix + "WarnMappingLoss";
        private const string AutoSaveEnabled =
            Prefix + "AutoSaveEnabled";
        private const string AutoSaveMinutes =
            Prefix + "AutoSaveMinutes";

        private const string UnitRateEnvironmentProperty =
            "TTBMVN.EstimateV2.UnitRateEnvironment";
        private const string OutputMapPrefix =
            "TTBMVN.EstimateV2.OutputMap.";

        private static readonly EstimateV2OutputSlot[] Slots =
        {
            EstimateV2OutputSlot.CostSummary,
            EstimateV2OutputSlot.EstimateAppendix,
            EstimateV2OutputSlot.UnitRateLand,
            EstimateV2OutputSlot.UnitRateWater,
            EstimateV2OutputSlot.ResourcePrices,
            EstimateV2OutputSlot.UnitRateSea
        };

        private static readonly WorksheetRole[] ManagedRoles =
        {
            WorksheetRole.CostSummary,
            WorksheetRole.EstimateAppendix,
            WorksheetRole.UnitRateLand,
            WorksheetRole.UnitRateWater,
            WorksheetRole.ResourcePrices
        };

        public static EstimateV2Settings Load(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            EstimateV2Settings defaults =
                EstimateV2SettingsPolicy.Defaults();

            var loaded = new EstimateV2Settings(
                ReadBoolProperty(
                    workbook,
                    AutoRestoreNorm,
                    defaults.AutoRestoreNormDisplay),
                ReadBoolProperty(
                    workbook,
                    ValidateOnOpen,
                    defaults.ValidateOnOpen),
                ReadBoolProperty(
                    workbook,
                    FormulaLinks,
                    defaults.FormulaLinksRequired),
                ReadBoolProperty(
                    workbook,
                    AutoSyncRows,
                    defaults.AutoSyncRows),
                ReadBoolProperty(
                    workbook,
                    UseCustomXml,
                    defaults.UseCustomXml),
                ReadBoolProperty(
                    workbook,
                    HideTechnical,
                    defaults.HideTechnicalColumns),
                ReadBoolProperty(
                    workbook,
                    WarnMappingLoss,
                    defaults.WarnOnMappingLoss),
                ReadBoolProperty(
                    workbook,
                    AutoSaveEnabled,
                    defaults.AutoSaveEnabled),
                ReadIntProperty(
                    workbook,
                    AutoSaveMinutes,
                    defaults.AutoSaveMinutes));

            return EstimateV2SettingsPolicy.Normalize(loaded);
        }

        public static void Save(
            Excel.Workbook workbook,
            EstimateV2Settings settings)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            EstimateV2Settings normalized =
                EstimateV2SettingsPolicy.Normalize(settings);

            object properties = null;
            try
            {
                properties =
                    workbook.CustomDocumentProperties;
                SetProperty(
                    properties,
                    AutoRestoreNorm,
                    normalized.AutoRestoreNormDisplay
                        ? "1"
                        : "0");
                SetProperty(
                    properties,
                    ValidateOnOpen,
                    normalized.ValidateOnOpen
                        ? "1"
                        : "0");
                SetProperty(
                    properties,
                    FormulaLinks,
                    "1");
                SetProperty(
                    properties,
                    AutoSyncRows,
                    normalized.AutoSyncRows
                        ? "1"
                        : "0");
                SetProperty(
                    properties,
                    UseCustomXml,
                    "1");
                SetProperty(
                    properties,
                    HideTechnical,
                    "1");
                SetProperty(
                    properties,
                    WarnMappingLoss,
                    normalized.WarnOnMappingLoss
                        ? "1"
                        : "0");
                SetProperty(
                    properties,
                    AutoSaveEnabled,
                    normalized.AutoSaveEnabled
                        ? "1"
                        : "0");
                SetProperty(
                    properties,
                    AutoSaveMinutes,
                    normalized.AutoSaveMinutes.ToString(
                        CultureInfo.InvariantCulture));
            }
            finally
            {
                Release(properties);
            }
        }

        public static void Reset(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            object properties = null;
            try
            {
                properties =
                    workbook.CustomDocumentProperties;
                dynamic items = properties;
                for (int index = items.Count;
                    index >= 1;
                    index--)
                {
                    object propertyObject = null;
                    try
                    {
                        dynamic property = items[index];
                        propertyObject = property;
                        string name =
                            Convert.ToString(
                                property.Name,
                                CultureInfo.InvariantCulture) ??
                            string.Empty;
                        if (name.StartsWith(
                            Prefix,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            property.Delete();
                        }
                    }
                    finally
                    {
                        Release(propertyObject);
                    }
                }
            }
            finally
            {
                Release(properties);
            }
        }

        public static WorkbookEstimateV2SettingsSnapshot
            CaptureSnapshot(
                Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            IReadOnlyList<WorkbookSheetDescriptor> sheets =
                WorkbookSheetChangeCoordinator
                    .CaptureSnapshot(workbook);
            var mappings =
                new Dictionary<EstimateV2OutputSlot, string>();
            int mappingLossCount = 0;

            foreach (EstimateV2OutputSlot slot in Slots)
            {
                string stored =
                    ReadOutputMap(
                        workbook,
                        slot);
                string resolved =
                    ResolveMappedSheetKey(
                        workbook,
                        slot);

                if (stored.Length > 0)
                {
                    mappings[slot] = stored;
                    if (!MappingIdentityMatches(
                        workbook,
                        slot,
                        stored))
                    {
                        mappingLossCount++;
                    }
                }
                else
                {
                    mappings[slot] = resolved;
                }
            }

            return new WorkbookEstimateV2SettingsSnapshot(
                Load(workbook),
                sheets,
                mappings,
                IsMetadataValid(workbook),
                mappingLossCount);
        }

        public static void SaveConfiguration(
            Excel.Workbook workbook,
            EstimateV2Settings settings,
            IReadOnlyDictionary<EstimateV2OutputSlot, string>
                selectedSheetKeys)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var requested =
                new Dictionary<EstimateV2OutputSlot, string>();
            foreach (EstimateV2OutputSlot slot in Slots)
            {
                string value = string.Empty;
                if (selectedSheetKeys != null)
                    selectedSheetKeys.TryGetValue(slot, out value);
                requested[slot] =
                    (value ?? string.Empty).Trim();
            }

            string[] duplicates = requested.Values
                .Where(value => value.Length > 0)
                .GroupBy(
                    value => value,
                    StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToArray();
            if (duplicates.Length > 0)
            {
                throw new InvalidOperationException(
                    "Một sheet không được gán cho nhiều đầu ra Dự toán V2.");
            }

            IReadOnlyList<WorkbookSheetDescriptor> available =
                WorkbookSheetChangeCoordinator
                    .CaptureSnapshot(workbook);
            var keys =
                new HashSet<string>(
                    available.Select(item => item.Key),
                    StringComparer.OrdinalIgnoreCase);
            foreach (string key in requested.Values
                .Where(value => value.Length > 0))
            {
                if (!keys.Contains(key))
                {
                    throw new InvalidOperationException(
                        "Sheet đã chọn không còn tồn tại: " +
                        key + ".");
                }
            }

            IReadOnlyList<WorksheetRoleAssignment>
                previousRoles =
                    WorksheetRoleService
                        .ReadAssignments(workbook);
            string[] previousSeaKeys =
                CaptureSeaSheetKeys(workbook);
            EstimateV2Settings previousSettings =
                Load(workbook);
            var previousOutputMaps =
                Slots.ToDictionary(
                    slot => slot,
                    slot => ReadOutputMap(
                        workbook,
                        slot));

            using (new ExcelWriteContext(
                workbook.Application))
            {
                try
                {
                    ClearManagedMappings(workbook);

                    foreach (EstimateV2OutputSlot slot in Slots)
                    {
                        string key = requested[slot];
                        if (key.Length == 0)
                            continue;

                        Excel.Worksheet sheet =
                            FindWorksheetByKey(
                                workbook,
                                key);
                        if (sheet == null)
                        {
                            throw new InvalidOperationException(
                                "Không tìm thấy sheet đã chọn: " +
                                key + ".");
                        }

                        try
                        {
                            WorksheetRole? role =
                                RoleForSlot(slot);
                            if (role.HasValue)
                            {
                                WorksheetRoleService.SetRole(
                                    sheet,
                                    role.Value);
                            }
                            else if (slot ==
                                EstimateV2OutputSlot.UnitRateSea)
                            {
                                SetWorksheetProperty(
                                    sheet,
                                    UnitRateEnvironmentProperty,
                                    EstimateV2RateEnvironment
                                        .Sea.ToString());
                            }
                        }
                        finally
                        {
                            Release(sheet);
                        }
                    }

                    foreach (EstimateV2OutputSlot slot in Slots)
                    {
                        WriteOutputMap(
                            workbook,
                            slot,
                            requested[slot]);
                    }

                    Save(workbook, settings);
                }
                catch
                {
                    try
                    {
                        WorksheetRoleService.RestoreAssignments(
                            workbook,
                            previousRoles);
                        ClearSeaMappings(workbook);
                        foreach (string key in previousSeaKeys)
                        {
                            Excel.Worksheet oldSea =
                                FindWorksheetByKey(
                                    workbook,
                                    key);
                            if (oldSea == null)
                                continue;
                            try
                            {
                                SetWorksheetProperty(
                                    oldSea,
                                    UnitRateEnvironmentProperty,
                                    EstimateV2RateEnvironment
                                        .Sea.ToString());
                            }
                            finally
                            {
                                Release(oldSea);
                            }
                        }

                        foreach (EstimateV2OutputSlot slot in Slots)
                        {
                            WriteOutputMap(
                                workbook,
                                slot,
                                previousOutputMaps[slot]);
                        }

                        Save(
                            workbook,
                            previousSettings);
                    }
                    catch (Exception rollbackException)
                    {
                        RuntimeLogger.Log(
                            rollbackException,
                            "Rollback Estimate V2 settings mapping");
                    }
                    throw;
                }
            }
        }

        public static int CountMappingLoss(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            int count = 0;
            foreach (EstimateV2OutputSlot slot in Slots)
            {
                string stored =
                    ReadOutputMap(
                        workbook,
                        slot);
                if (stored.Length > 0 &&
                    !MappingIdentityMatches(
                        workbook,
                        slot,
                        stored))
                {
                    count++;
                }
            }
            return count;
        }

        private static bool MappingIdentityMatches(
            Excel.Workbook workbook,
            EstimateV2OutputSlot slot,
            string sheetKey)
        {
            Excel.Worksheet sheet = null;
            try
            {
                sheet = FindWorksheetByKey(
                    workbook,
                    sheetKey);
                if (sheet == null)
                    return false;

                WorksheetRole? role =
                    RoleForSlot(slot);
                if (role.HasValue)
                {
                    WorksheetRole actual;
                    return WorksheetRoleService.TryGetRole(
                        sheet,
                        out actual) &&
                        actual == role.Value;
                }

                if (slot ==
                    EstimateV2OutputSlot.UnitRateSea)
                {
                    return string.Equals(
                        ReadWorksheetProperty(
                            sheet,
                            UnitRateEnvironmentProperty),
                        EstimateV2RateEnvironment.Sea
                            .ToString(),
                        StringComparison.OrdinalIgnoreCase);
                }

                return false;
            }
            finally
            {
                Release(sheet);
            }
        }

        private static string ReadOutputMap(
            Excel.Workbook workbook,
            EstimateV2OutputSlot slot)
        {
            string value;
            return TryReadProperty(
                    workbook,
                    OutputMapPrefix +
                    slot.ToString(),
                    out value)
                ? (value ?? string.Empty).Trim()
                : string.Empty;
        }

        private static void WriteOutputMap(
            Excel.Workbook workbook,
            EstimateV2OutputSlot slot,
            string sheetKey)
        {
            string name =
                OutputMapPrefix +
                slot.ToString();
            string value =
                (sheetKey ?? string.Empty)
                    .Trim();

            object properties = null;
            try
            {
                properties =
                    workbook.CustomDocumentProperties;
                if (value.Length == 0)
                {
                    DeleteProperty(
                        properties,
                        name);
                }
                else
                {
                    SetProperty(
                        properties,
                        name,
                        value);
                }
            }
            finally
            {
                Release(properties);
            }
        }

        private static void ClearManagedMappings(
            Excel.Workbook workbook)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;

                        WorksheetRole role;
                        if (WorksheetRoleService.TryGetRole(
                            sheet,
                            out role) &&
                            ManagedRoles.Contains(role))
                        {
                            WorksheetRoleService.ClearRole(
                                sheet);
                        }

                        string environment =
                            ReadWorksheetProperty(
                                sheet,
                                UnitRateEnvironmentProperty);
                        if (string.Equals(
                            environment,
                            EstimateV2RateEnvironment.Sea
                                .ToString(),
                            StringComparison.OrdinalIgnoreCase))
                        {
                            DeleteWorksheetProperty(
                                sheet,
                                UnitRateEnvironmentProperty);
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
            }
            finally
            {
                Release(sheets);
            }
        }

        private static void ClearSeaMappings(
            Excel.Workbook workbook)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;
                        string environment =
                            ReadWorksheetProperty(
                                sheet,
                                UnitRateEnvironmentProperty);
                        if (string.Equals(
                            environment,
                            EstimateV2RateEnvironment.Sea
                                .ToString(),
                            StringComparison.OrdinalIgnoreCase))
                        {
                            DeleteWorksheetProperty(
                                sheet,
                                UnitRateEnvironmentProperty);
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
            }
            finally
            {
                Release(sheets);
            }
        }

        private static string ResolveMappedSheetKey(
            Excel.Workbook workbook,
            EstimateV2OutputSlot slot)
        {
            Excel.Worksheet sheet = null;
            try
            {
                WorksheetRole? role =
                    RoleForSlot(slot);
                if (role.HasValue)
                {
                    sheet =
                        WorkbookEstimateV2CompatibilityService
                            .ResolveOutputWorksheet(
                                workbook,
                                role.Value,
                                CanonicalNames(slot));
                }
                else
                {
                    sheet =
                        WorkbookEstimateV2CompatibilityService
                            .ResolveOutputWorksheetByKind(
                                workbook,
                                EstimateV2LegacySheetKind
                                    .UnitRateSea,
                                CanonicalNames(slot));
                }

                if (sheet == null)
                    return string.Empty;

                string key =
                    (sheet.CodeName ??
                        string.Empty).Trim();
                return key.Length > 0
                    ? key
                    : (sheet.Name ??
                        string.Empty).Trim();
            }
            finally
            {
                Release(sheet);
            }
        }

        private static string[] CanonicalNames(
            EstimateV2OutputSlot slot)
        {
            switch (slot)
            {
                case EstimateV2OutputSlot.CostSummary:
                    return new[] { "THKP-TC" };
                case EstimateV2OutputSlot.EstimateAppendix:
                    return new[] { "Gia DT TC" };
                case EstimateV2OutputSlot.UnitRateLand:
                    return new[] { "DG Can", "DG Cạn" };
                case EstimateV2OutputSlot.UnitRateWater:
                    return new[] { "DG Nuoc", "DG Nước" };
                case EstimateV2OutputSlot.ResourcePrices:
                    return new[] { "VL-NC-M" };
                case EstimateV2OutputSlot.UnitRateSea:
                    return new[] { "DG Bien", "DG Biển" };
                default:
                    return new string[0];
            }
        }

        private static WorksheetRole? RoleForSlot(
            EstimateV2OutputSlot slot)
        {
            switch (slot)
            {
                case EstimateV2OutputSlot.CostSummary:
                    return WorksheetRole.CostSummary;
                case EstimateV2OutputSlot.EstimateAppendix:
                    return WorksheetRole.EstimateAppendix;
                case EstimateV2OutputSlot.UnitRateLand:
                    return WorksheetRole.UnitRateLand;
                case EstimateV2OutputSlot.UnitRateWater:
                    return WorksheetRole.UnitRateWater;
                case EstimateV2OutputSlot.ResourcePrices:
                    return WorksheetRole.ResourcePrices;
                default:
                    return null;
            }
        }

        private static bool IsMetadataValid(
            Excel.Workbook workbook)
        {
            try
            {
                IReadOnlyList<EstimateV2RegisteredSource> sources =
                    WorkbookEstimateV2RegistrationService
                        .ListRegistered(workbook);
                if (sources.Count == 0)
                    return false;

                EstimateV2State state;
                if (!WorkbookEstimateV2StateService.TryLoad(
                    workbook,
                    out state) ||
                    state == null)
                {
                    return false;
                }

                return sources.All(source =>
                    source.Columns.TechnicalIdColumn >
                        source.Columns.QuantityColumn &&
                    source.Columns.TechnicalNormColumn ==
                        source.Columns.TechnicalIdColumn + 1 &&
                    source.Columns.TechnicalKindColumn ==
                        source.Columns.TechnicalIdColumn + 2 &&
                    source.Columns
                        .TechnicalFingerprintColumn ==
                        source.Columns.TechnicalIdColumn + 3);
            }
            catch
            {
                return false;
            }
        }

        private static string[] CaptureSeaSheetKeys(
            Excel.Workbook workbook)
        {
            var keys = new List<string>();
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;
                        string environment =
                            ReadWorksheetProperty(
                                sheet,
                                UnitRateEnvironmentProperty);
                        if (!string.Equals(
                            environment,
                            EstimateV2RateEnvironment.Sea
                                .ToString(),
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string key =
                            (sheet.CodeName ??
                                string.Empty).Trim();
                        if (key.Length == 0)
                            key =
                                (sheet.Name ??
                                    string.Empty).Trim();
                        if (key.Length > 0)
                            keys.Add(key);
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
            }
            finally
            {
                Release(sheets);
            }
            return keys.ToArray();
        }

        private static Excel.Worksheet FindWorksheetByKey(
            Excel.Workbook workbook,
            string key)
        {
            string wanted =
                (key ?? string.Empty).Trim();
            if (wanted.Length == 0)
                return null;

            Excel.Sheets sheets = null;
            Excel.Worksheet nameFallback = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;

                        string codeName =
                            (sheet.CodeName ??
                                string.Empty).Trim();
                        if (codeName.Length > 0 &&
                            string.Equals(
                                codeName,
                                wanted,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Worksheet result =
                                sheet;
                            sheet = null;
                            Release(nameFallback);
                            return result;
                        }

                        if (nameFallback == null &&
                            string.Equals(
                                sheet.Name,
                                wanted,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            nameFallback = sheet;
                            sheet = null;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }

                return nameFallback;
            }
            finally
            {
                Release(sheets);
            }
        }

        private static bool ReadBoolProperty(
            Excel.Workbook workbook,
            string name,
            bool defaultValue)
        {
            string value;
            if (!TryReadProperty(
                workbook,
                name,
                out value))
            {
                return defaultValue;
            }

            if (string.Equals(
                value,
                "1",
                StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    value,
                    "true",
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(
                value,
                "0",
                StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    value,
                    "false",
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return defaultValue;
        }

        private static int ReadIntProperty(
            Excel.Workbook workbook,
            string name,
            int defaultValue)
        {
            string value;
            int parsed;
            return TryReadProperty(
                    workbook,
                    name,
                    out value) &&
                int.TryParse(
                    value,
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out parsed)
                ? parsed
                : defaultValue;
        }

        private static bool TryReadProperty(
            Excel.Workbook workbook,
            string name,
            out string value)
        {
            object properties = null;
            try
            {
                properties =
                    workbook.CustomDocumentProperties;
                dynamic items = properties;
                for (int index = 1;
                    index <= items.Count;
                    index++)
                {
                    object propertyObject = null;
                    try
                    {
                        dynamic property = items[index];
                        propertyObject = property;
                        if (!string.Equals(
                            Convert.ToString(
                                property.Name,
                                CultureInfo.InvariantCulture),
                            name,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        value =
                            Convert.ToString(
                                property.Value,
                                CultureInfo.InvariantCulture) ??
                            string.Empty;
                        return true;
                    }
                    finally
                    {
                        Release(propertyObject);
                    }
                }
            }
            catch (COMException)
            {
            }
            finally
            {
                Release(properties);
            }

            value = null;
            return false;
        }

        private static void DeleteProperty(
            object propertiesObject,
            string name)
        {
            dynamic properties = propertiesObject;
            for (int index = properties.Count;
                index >= 1;
                index--)
            {
                object propertyObject = null;
                try
                {
                    dynamic property = properties[index];
                    propertyObject = property;
                    if (string.Equals(
                        Convert.ToString(
                            property.Name,
                            CultureInfo.InvariantCulture),
                        name,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        property.Delete();
                    }
                }
                finally
                {
                    Release(propertyObject);
                }
            }
        }

        private static void SetProperty(
            object propertiesObject,
            string name,
            string value)
        {
            dynamic properties = propertiesObject;
            object existingObject = null;
            try
            {
                for (int index = 1;
                    index <= properties.Count;
                    index++)
                {
                    object propertyObject = null;
                    try
                    {
                        dynamic property = properties[index];
                        propertyObject = property;
                        if (!string.Equals(
                            Convert.ToString(
                                property.Name,
                                CultureInfo.InvariantCulture),
                            name,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        existingObject =
                            propertyObject;
                        propertyObject = null;
                        break;
                    }
                    finally
                    {
                        Release(propertyObject);
                    }
                }

                if (existingObject == null)
                {
                    existingObject =
                        properties.Add(
                            name,
                            false,
                            MsoDocProperties
                                .msoPropertyTypeString,
                            value,
                            Type.Missing);
                }
                else
                {
                    dynamic existing =
                        existingObject;
                    existing.Value =
                        value ?? string.Empty;
                }
            }
            finally
            {
                Release(existingObject);
            }
        }

        private static string ReadWorksheetProperty(
            Excel.Worksheet sheet,
            string name)
        {
            Excel.CustomProperties properties = null;
            try
            {
                properties =
                    sheet.CustomProperties;
                for (int index = 1;
                    index <= properties.Count;
                    index++)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property =
                            properties.Item[index];
                        if (string.Equals(
                            property.Name,
                            name,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            return Convert.ToString(
                                property.Value,
                                CultureInfo.InvariantCulture) ??
                                string.Empty;
                        }
                    }
                    finally
                    {
                        Release(property);
                    }
                }
                return string.Empty;
            }
            finally
            {
                Release(properties);
            }
        }

        private static void SetWorksheetProperty(
            Excel.Worksheet sheet,
            string name,
            string value)
        {
            Excel.CustomProperties properties = null;
            Excel.CustomProperty first = null;
            try
            {
                properties =
                    sheet.CustomProperties;
                for (int index = properties.Count;
                    index >= 1;
                    index--)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property =
                            properties.Item[index];
                        if (!string.Equals(
                            property.Name,
                            name,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (first == null)
                        {
                            first = property;
                            property = null;
                        }
                        else
                        {
                            property.Delete();
                        }
                    }
                    finally
                    {
                        Release(property);
                    }
                }

                if (first == null)
                    first = properties.Add(name, value);
                else
                    first.Value = value;
            }
            finally
            {
                Release(first);
                Release(properties);
            }
        }

        private static void DeleteWorksheetProperty(
            Excel.Worksheet sheet,
            string name)
        {
            Excel.CustomProperties properties = null;
            try
            {
                properties =
                    sheet.CustomProperties;
                for (int index = properties.Count;
                    index >= 1;
                    index--)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property =
                            properties.Item[index];
                        if (string.Equals(
                            property.Name,
                            name,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            property.Delete();
                        }
                    }
                    finally
                    {
                        Release(property);
                    }
                }
            }
            finally
            {
                Release(properties);
            }
        }

        private static void Release(
            object value)
        {
            if (value != null &&
                Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
