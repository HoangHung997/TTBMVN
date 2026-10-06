using System;

namespace ExcelAddIn1.Core
{
    public sealed class EstimateV2Settings
    {
        public EstimateV2Settings(
            bool autoRestoreNormDisplay,
            bool validateOnOpen,
            bool formulaLinksRequired,
            bool autoSyncRows,
            bool useCustomXml,
            bool hideTechnicalColumns,
            bool warnOnMappingLoss,
            bool autoSaveEnabled,
            int autoSaveMinutes)
        {
            AutoRestoreNormDisplay = autoRestoreNormDisplay;
            ValidateOnOpen = validateOnOpen;
            FormulaLinksRequired = formulaLinksRequired;
            AutoSyncRows = autoSyncRows;
            UseCustomXml = useCustomXml;
            HideTechnicalColumns = hideTechnicalColumns;
            WarnOnMappingLoss = warnOnMappingLoss;
            AutoSaveEnabled = autoSaveEnabled;
            AutoSaveMinutes = autoSaveMinutes;
        }

        public bool AutoRestoreNormDisplay { get; }
        public bool ValidateOnOpen { get; }
        public bool FormulaLinksRequired { get; }
        public bool AutoSyncRows { get; }
        public bool UseCustomXml { get; }
        public bool HideTechnicalColumns { get; }
        public bool WarnOnMappingLoss { get; }
        public bool AutoSaveEnabled { get; }
        public int AutoSaveMinutes { get; }
    }

    public static class EstimateV2SettingsPolicy
    {
        public const int DefaultAutoSaveMinutes = 5;
        public const int MinimumAutoSaveMinutes = 1;
        public const int MaximumAutoSaveMinutes = 60;

        public static EstimateV2Settings Defaults()
        {
            return new EstimateV2Settings(
                true,
                true,
                true,
                true,
                true,
                true,
                true,
                true,
                DefaultAutoSaveMinutes);
        }

        public static EstimateV2Settings Normalize(
            EstimateV2Settings settings)
        {
            EstimateV2Settings source =
                settings ?? Defaults();

            int minutes = source.AutoSaveMinutes;
            if (minutes < MinimumAutoSaveMinutes)
                minutes = MinimumAutoSaveMinutes;
            if (minutes > MaximumAutoSaveMinutes)
                minutes = MaximumAutoSaveMinutes;

            // Ba quy tac nay la invariant cua Du toan V2:
            // - ket qua tinh la formula/link, khong so chet;
            // - binding that nam trong Custom XML;
            // - metadata ky thuat khong hien trong ho so in.
            // UI co the hien trang thai nhung khong cho phep tat.
            return new EstimateV2Settings(
                source.AutoRestoreNormDisplay,
                source.ValidateOnOpen,
                true,
                source.AutoSyncRows,
                true,
                true,
                source.WarnOnMappingLoss,
                source.AutoSaveEnabled,
                minutes);
        }
    }
}
