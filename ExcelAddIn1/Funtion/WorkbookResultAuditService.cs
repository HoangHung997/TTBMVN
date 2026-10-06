using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookResultAuditService
    {
        public const string ManifestPropertyName = "TTBMVN.ResultAudit.Manifest";
        private const string PartPropertyPrefix = "TTBMVN.ResultAudit.Part.";
        public const string EstimateScopeId = "EstimateAppendix";
        public const string CostSummaryScopeId = "CostSummary";
        private const string PlanMetadataPrefix = "TTBMVN_PLAN|";
        private const string BindingMetadataPrefix = "TTBMVN_BINDING|";
        private const string BlockMetadataPrefix = "TTBMVN_BLOCK|";

        public static ResultAuditTrail Load(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            string payload;
            if (!WorkbookCustomPayloadStore.TryRead(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                "Result audit",
                out payload))
            {
                return new ResultAuditTrail(Enumerable.Empty<ResultAuditEntry>());
            }
            return ResultAuditTrailSerializer.Deserialize(payload);
        }

        public static bool SaveScope(
            Excel.Workbook workbook,
            string scopeId,
            IEnumerable<ResultAuditEntry> entries)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            ResultAuditTrail updated = Load(workbook).ReplaceScope(scopeId, entries);
            return WorkbookCustomPayloadStore.Save(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                ResultAuditTrailSerializer.Serialize(updated),
                "Result audit");
        }

        internal static bool TryReadPayload(Excel.Workbook workbook, out string payload)
        {
            return WorkbookCustomPayloadStore.TryRead(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                "Result audit",
                out payload);
        }

        internal static bool RestorePayload(Excel.Workbook workbook, string payload)
        {
            if (payload == null)
            {
                return WorkbookCustomPayloadStore.Clear(
                    workbook,
                    ManifestPropertyName,
                    PartPropertyPrefix,
                    "Result audit");
            }

            ResultAuditTrailSerializer.Deserialize(payload);
            return WorkbookCustomPayloadStore.Save(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                payload,
                "Result audit");
        }

        public static ResultAuditEntry Find(
            Excel.Workbook workbook,
            string worksheetCodeName,
            int row,
            int column)
        {
            return Load(workbook).FindMostSpecific(worksheetCodeName, row, column);
        }

        public static ResultAuditEntry FindAtActiveCell(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            Excel.Range activeCell = null;
            Excel.Worksheet worksheet = null;
            Excel.Workbook selectedWorkbook = null;
            try
            {
                activeCell = workbook.Application.ActiveCell as Excel.Range;
                if (activeCell == null)
                    throw new InvalidOperationException("Khong co o Excel dang chon.");
                worksheet = activeCell.Worksheet as Excel.Worksheet;
                selectedWorkbook = worksheet?.Parent as Excel.Workbook;
                if (worksheet == null || selectedWorkbook == null ||
                    !string.Equals(selectedWorkbook.FullName, workbook.FullName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("O dang chon khong thuoc workbook du toan nay.");
                }
                return Find(workbook, worksheet.CodeName, activeCell.Row, activeCell.Column);
            }
            finally
            {
                Release(selectedWorkbook);
                Release(worksheet);
                Release(activeCell);
            }
        }

        public static IReadOnlyList<ResultAuditEntry> BuildEstimateEntries(
            WorkbookEstimatePreview preview,
            string worksheetCodeName,
            int firstRow,
            int lastRow)
        {
            if (preview == null)
                throw new ArgumentNullException(nameof(preview));
            if (!preview.IsValid || preview.Result == null)
                throw new InvalidOperationException("Preview phu luc khong hop le de tao audit.");

            EstimateAppendixCalculationResult result = preview.Result;
            var entries = new List<ResultAuditEntry>();
            var resultLines = result.Lines.ToDictionary(
                line => line.Request.LineId,
                StringComparer.OrdinalIgnoreCase);
            foreach (WorkbookEstimateLine planLine in preview.Plan.Lines)
            {
                EstimateAppendixLineResult resultLine;
                if (!resultLines.TryGetValue(planLine.LineId, out resultLine))
                    continue;
                UnitRateCalculationResult rate = resultLine.Request.UnitRate;
                var sources = new List<ResultAuditSource>();
                AddRegulationSource(sources, rate?.NormSource);
                if (rate != null)
                {
                    foreach (string reference in rate.Resources
                        .Where(resource => !resource.IsPercentage)
                        .Select(resource => resource.PriceSourceReference)
                        .Where(reference => !string.IsNullOrWhiteSpace(reference))
                        .Distinct(StringComparer.Ordinal)
                        .OrderBy(reference => reference, StringComparer.Ordinal))
                    {
                        sources.Add(ReferenceSource(ResultAuditSourceKind.Price, reference));
                    }
                }
                AddReference(sources, ResultAuditSourceKind.Quantity, planLine.QuantitySourceReference);
                AddReference(sources, ResultAuditSourceKind.Output, planLine.OutputReference);
                sources.Add(ReferenceSource(
                    ResultAuditSourceKind.Metadata,
                    BuildPlanMetadata(planLine)));
                foreach (UnitRateResourceBinding binding in planLine.Bindings)
                {
                    sources.Add(ReferenceSource(
                        ResultAuditSourceKind.Metadata,
                        BuildBindingMetadata(binding)));
                }
                entries.Add(new ResultAuditEntry(
                    "EST|" + worksheetCodeName + "|" + planLine.TargetRow.ToString(CultureInfo.InvariantCulture),
                    EstimateScopeId,
                    ResultAuditKind.EstimateLine,
                    worksheetCodeName,
                    WorksheetRoleCatalog.ToId(WorksheetRole.EstimateAppendix),
                    planLine.TargetRow,
                    6,
                    planLine.TargetRow,
                    14,
                    planLine.Description.Length == 0 ? planLine.LineId : planLine.Description,
                    result.PackageId,
                    result.PackageVersion,
                    result.PackageChecksum,
                    result.UnitRateProfileId,
                    result.UnitRateProfileVersion,
                    result.UnitRateProfileChecksum,
                    planLine.NormKey,
                    planLine.VariantCode,
                    DistinctSources(sources)));
            }

            var blockSources = entries
                .SelectMany(entry => entry.Sources)
                .Where(source => source.Kind != ResultAuditSourceKind.Metadata)
                .ToList();
            blockSources.Add(ReferenceSource(
                ResultAuditSourceKind.Metadata,
                BlockMetadataPrefix +
                    preview.Plan.AdjustmentRow.ToString(CultureInfo.InvariantCulture) + "|" +
                    preview.Plan.GrandTotalRow.ToString(CultureInfo.InvariantCulture) + "|" +
                    Encode(string.Join(";", preview.Conditions))));
            entries.Add(new ResultAuditEntry(
                "EST|" + worksheetCodeName + "|BLOCK",
                EstimateScopeId,
                ResultAuditKind.EstimateBlock,
                worksheetCodeName,
                WorksheetRoleCatalog.ToId(WorksheetRole.EstimateAppendix),
                firstRow,
                6,
                lastRow,
                14,
                "Phu luc du toan thi cong",
                result.PackageId,
                result.PackageVersion,
                result.PackageChecksum,
                result.UnitRateProfileId,
                result.UnitRateProfileVersion,
                result.UnitRateProfileChecksum,
                string.Empty,
                string.Empty,
                DistinctSources(blockSources)));
            return entries.AsReadOnly();
        }

        public static bool TryRestoreEstimatePlan(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            WorkbookUnitRateContext context,
            out WorkbookEstimatePlan plan)
        {
            return TryRestoreEstimatePlan(
                workbook,
                worksheet,
                context,
                true,
                out plan);
        }

        internal static bool TryRestoreEstimatePlanForValidation(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            WorkbookUnitRateContext context,
            out WorkbookEstimatePlan plan)
        {
            return TryRestoreEstimatePlan(
                workbook,
                worksheet,
                context,
                false,
                out plan);
        }

        private static bool TryRestoreEstimatePlan(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            WorkbookUnitRateContext context,
            bool validateIdentity,
            out WorkbookEstimatePlan plan)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            ResultAuditEntry[] entries = Load(workbook).Entries
                .Where(entry => entry.Kind == ResultAuditKind.EstimateLine &&
                    string.Equals(entry.ScopeId, EstimateScopeId, StringComparison.Ordinal) &&
                    string.Equals(entry.WorksheetCodeName, worksheet.CodeName, StringComparison.Ordinal))
                .OrderBy(entry => entry.FirstRow)
                .ToArray();
            if (entries.Length == 0)
            {
                plan = null;
                return false;
            }
            if (validateIdentity)
                ValidateEstimateIdentity(entries, context);

            int firstRow = entries.Min(entry => entry.FirstRow);
            int lastRow = entries.Max(entry => entry.LastRow);
            Excel.Range dataRange = null;
            try
            {
                dataRange = worksheet.Range[
                    "B" + firstRow.ToString(CultureInfo.InvariantCulture),
                    "E" + lastRow.ToString(CultureInfo.InvariantCulture)];
                object values = dataRange.Value2;
                var lines = new List<WorkbookEstimateLine>();
                foreach (ResultAuditEntry entry in entries)
                {
                    PlanMetadata metadata = ParsePlanMetadata(entry);
                    int offset = entry.FirstRow - firstRow + 1;
                    lines.Add(new WorkbookEstimateLine(
                        metadata.LineId,
                        metadata.GroupKey,
                        entry.FirstRow,
                        Convert.ToString(ArrayValue(values, offset, 1), CultureInfo.CurrentCulture)?.Trim() ?? entry.Label,
                        Convert.ToString(ArrayValue(values, offset, 2), CultureInfo.CurrentCulture)?.Trim() ?? string.Empty,
                        ToDecimal(ArrayValue(values, offset, 3), "D" + entry.FirstRow),
                        ToDecimal(ArrayValue(values, offset, 4), "E" + entry.FirstRow),
                        SourceReference(entry, ResultAuditSourceKind.Quantity),
                        SourceReference(entry, ResultAuditSourceKind.Output),
                        entry.NormKey,
                        entry.VariantCode,
                        metadata.ImportWarning,
                        metadata.ImportNote,
                        metadata.UnitRateRole,
                        metadata.UnitRateTotalRow,
                        ParseBindings(entry)));
                }

                int adjustmentRow;
                int grandTotalRow;
                ParseBlockMetadata(
                    Load(workbook).Entries.FirstOrDefault(entry =>
                        entry.Kind == ResultAuditKind.EstimateBlock &&
                        string.Equals(entry.WorksheetCodeName, worksheet.CodeName, StringComparison.Ordinal)),
                    out adjustmentRow,
                    out grandTotalRow);
                plan = new WorkbookEstimatePlan(
                    worksheet.CodeName,
                    lines,
                    adjustmentRow,
                    grandTotalRow);
                return true;
            }
            finally
            {
                Release(dataRange);
            }
        }

        public static IReadOnlyList<string> ReadEstimateConditions(
            Excel.Workbook workbook,
            string worksheetCodeName)
        {
            ResultAuditEntry block = Load(workbook).Entries.FirstOrDefault(entry =>
                entry.Kind == ResultAuditKind.EstimateBlock &&
                string.Equals(entry.WorksheetCodeName, worksheetCodeName, StringComparison.Ordinal));
            string value = block?.Sources
                .Where(source => source.Kind == ResultAuditSourceKind.Metadata)
                .Select(source => source.Reference)
                .FirstOrDefault(reference => reference.StartsWith(BlockMetadataPrefix, StringComparison.Ordinal));
            string[] parts = (value ?? string.Empty).Split('|');
            if (parts.Length < 4)
                return new string[0];
            return Decode(parts[3])
                .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(condition => condition.Trim())
                .Where(condition => condition.Length > 0)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(condition => condition, StringComparer.Ordinal)
                .ToArray();
        }

        public static IReadOnlyList<ResultAuditEntry> BuildCostSummaryEntries(
            Excel.Workbook workbook,
            WorkbookCostSummaryPreview preview,
            string worksheetCodeName)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (preview == null)
                throw new ArgumentNullException(nameof(preview));

            PriceProfile profile = WorkbookPriceProfileService.LoadRequired(workbook);
            var entries = new List<ResultAuditEntry>();
            foreach (CostComponentCalculationResult component in preview.Result.Components)
            {
                int row = 17 + (component.ComponentCode[1] - '0');
                var sources = new List<ResultAuditSource>();
                AddRegulationSource(sources, component.Source);
                AddReference(sources, ResultAuditSourceKind.External, component.ExternalBasis);
                entries.Add(new ResultAuditEntry(
                    "THKP|" + worksheetCodeName + "|" + component.ComponentCode,
                    CostSummaryScopeId,
                    ResultAuditKind.CostComponent,
                    worksheetCodeName,
                    WorksheetRoleCatalog.ToId(WorksheetRole.CostSummary),
                    row,
                    4,
                    row,
                    5,
                    component.ComponentCode,
                    preview.Context.PackageId,
                    preview.Context.PackageVersion,
                    preview.Context.PackageChecksum,
                    profile.ProfileId,
                    profile.DataVersion,
                    profile.Checksum,
                    string.Empty,
                    string.Empty,
                    DistinctSources(sources)));
            }

            var summarySources = entries.SelectMany(entry => entry.Sources).ToList();
            AddReference(
                summarySources,
                ResultAuditSourceKind.Output,
                preview.Context.EstimateSheetCodeName + "!I27:K27");
            entries.Add(new ResultAuditEntry(
                "THKP|" + worksheetCodeName + "|BLOCK",
                CostSummaryScopeId,
                ResultAuditKind.CostSummary,
                worksheetCodeName,
                WorksheetRoleCatalog.ToId(WorksheetRole.CostSummary),
                10,
                4,
                27,
                5,
                "Tong hop kinh phi",
                preview.Context.PackageId,
                preview.Context.PackageVersion,
                preview.Context.PackageChecksum,
                profile.ProfileId,
                profile.DataVersion,
                profile.Checksum,
                string.Empty,
                string.Empty,
                DistinctSources(summarySources)));
            entries.Add(new ResultAuditEntry(
                "THKP|" + worksheetCodeName + "|WORDS",
                CostSummaryScopeId,
                ResultAuditKind.CostSummary,
                worksheetCodeName,
                WorksheetRoleCatalog.ToId(WorksheetRole.CostSummary),
                28,
                1,
                28,
                1,
                "Tong kinh phi bang chu",
                preview.Context.PackageId,
                preview.Context.PackageVersion,
                preview.Context.PackageChecksum,
                profile.ProfileId,
                profile.DataVersion,
                profile.Checksum,
                string.Empty,
                string.Empty,
                DistinctSources(summarySources)));
            return entries.AsReadOnly();
        }

        private static void AddRegulationSource(
            ICollection<ResultAuditSource> sources,
            RegulationSourceLocator locator)
        {
            if (locator == null)
                return;
            sources.Add(new ResultAuditSource(
                ResultAuditSourceKind.Regulation,
                locator.DocumentId,
                locator.PageFrom,
                locator.PageTo,
                locator.Section,
                string.Empty));
        }

        private static string BuildPlanMetadata(WorkbookEstimateLine line)
        {
            return PlanMetadataPrefix + string.Join("|", new[]
            {
                Encode(line.LineId),
                Encode(line.GroupKey),
                WorksheetRoleCatalog.ToId(line.UnitRateRole),
                line.UnitRateTotalRow.ToString(CultureInfo.InvariantCulture),
                Encode(line.ImportWarning),
                Encode(line.ImportNote)
            });
        }

        private static string BuildBindingMetadata(UnitRateResourceBinding binding)
        {
            return BindingMetadataPrefix + string.Join("|", new[]
            {
                Encode(binding.ResourceCode),
                Encode(binding.PriceCode),
                Encode(binding.Reason)
            });
        }

        private static PlanMetadata ParsePlanMetadata(ResultAuditEntry entry)
        {
            string value = entry.Sources
                .Where(source => source.Kind == ResultAuditSourceKind.Metadata)
                .Select(source => source.Reference)
                .FirstOrDefault(reference => reference.StartsWith(PlanMetadataPrefix, StringComparison.Ordinal));
            string[] parts = (value ?? string.Empty).Split('|');
            WorksheetRole role;
            int totalRow;
            if (parts.Length != 7 || !WorksheetRoleCatalog.TryParse(parts[3], out role) ||
                !int.TryParse(parts[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out totalRow) ||
                totalRow < 1)
            {
                throw new InvalidOperationException("Audit phu luc thieu metadata dong " + entry.EntryId + ".");
            }
            return new PlanMetadata(
                Decode(parts[1]), Decode(parts[2]), role, totalRow, Decode(parts[5]), Decode(parts[6]));
        }

        private static IReadOnlyList<UnitRateResourceBinding> ParseBindings(ResultAuditEntry entry)
        {
            var output = new List<UnitRateResourceBinding>();
            foreach (string value in entry.Sources
                .Where(source => source.Kind == ResultAuditSourceKind.Metadata)
                .Select(source => source.Reference)
                .Where(reference => reference.StartsWith(BindingMetadataPrefix, StringComparison.Ordinal)))
            {
                string[] parts = value.Split('|');
                if (parts.Length != 4)
                    throw new InvalidOperationException("Audit binding khong hop le: " + entry.EntryId + ".");
                output.Add(new UnitRateResourceBinding(
                    Decode(parts[1]), Decode(parts[2]), Decode(parts[3])));
            }
            return output.AsReadOnly();
        }

        private static void ParseBlockMetadata(
            ResultAuditEntry block,
            out int adjustmentRow,
            out int grandTotalRow)
        {
            string value = block?.Sources
                .Where(source => source.Kind == ResultAuditSourceKind.Metadata)
                .Select(source => source.Reference)
                .FirstOrDefault(reference => reference.StartsWith(BlockMetadataPrefix, StringComparison.Ordinal));
            string[] parts = (value ?? string.Empty).Split('|');
            if ((parts.Length != 3 && parts.Length != 4) ||
                !int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out adjustmentRow) ||
                !int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out grandTotalRow) ||
                adjustmentRow < 0 || grandTotalRow < 1)
            {
                throw new InvalidOperationException("Audit phu luc thieu metadata khoi tong.");
            }
        }

        private static void ValidateEstimateIdentity(
            IEnumerable<ResultAuditEntry> entries,
            WorkbookUnitRateContext context)
        {
            foreach (ResultAuditEntry entry in entries)
            {
                if (!string.Equals(entry.PackageId, context.PackageId, StringComparison.Ordinal) ||
                    !string.Equals(entry.PackageVersion, context.PackageVersion, StringComparison.Ordinal) ||
                    !string.Equals(entry.PackageChecksum, context.PackageChecksum, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(entry.PriceProfileId, context.PriceProfile.ProfileId, StringComparison.Ordinal) ||
                    !string.Equals(entry.PriceProfileVersion, context.PriceProfile.DataVersion, StringComparison.Ordinal) ||
                    !string.Equals(entry.PriceProfileChecksum, context.PriceProfile.Checksum, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Audit phu luc khong khop package/PriceProfile dang pin. Hay chay kiem tra sai lech.");
                }
            }
        }

        private static string SourceReference(ResultAuditEntry entry, ResultAuditSourceKind kind)
        {
            return entry.Sources
                .Where(source => source.Kind == kind)
                .Select(source => source.Reference)
                .FirstOrDefault(reference => reference.Length > 0) ?? string.Empty;
        }

        private static object ArrayValue(object values, int row, int column)
        {
            Array array = values as Array;
            if (array == null)
                return row == 1 && column == 1 ? values : null;
            return array.GetValue(array.GetLowerBound(0) + row - 1, array.GetLowerBound(1) + column - 1);
        }

        private static decimal ToDecimal(object value, string address)
        {
            if (value == null || string.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.CurrentCulture)))
                return 0m;
            try
            {
                return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Gia tri " + address + " khong phai so.", ex);
            }
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value ?? string.Empty));
            }
            catch (FormatException ex)
            {
                throw new InvalidOperationException("Audit metadata khong hop le.", ex);
            }
        }

        private static void AddReference(
            ICollection<ResultAuditSource> sources,
            ResultAuditSourceKind kind,
            string reference)
        {
            if (!string.IsNullOrWhiteSpace(reference))
                sources.Add(ReferenceSource(kind, reference));
        }

        private static ResultAuditSource ReferenceSource(
            ResultAuditSourceKind kind,
            string reference)
        {
            return new ResultAuditSource(kind, string.Empty, 0, 0, string.Empty, reference);
        }

        private static IEnumerable<ResultAuditSource> DistinctSources(
            IEnumerable<ResultAuditSource> sources)
        {
            return sources
                .GroupBy(source => source.SortKey, StringComparer.Ordinal)
                .Select(group => group.First())
                .OrderBy(source => source.SortKey, StringComparer.Ordinal)
                .ToArray();
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private sealed class PlanMetadata
        {
            public PlanMetadata(
                string lineId,
                string groupKey,
                WorksheetRole unitRateRole,
                int unitRateTotalRow,
                string importWarning,
                string importNote)
            {
                LineId = lineId;
                GroupKey = groupKey;
                UnitRateRole = unitRateRole;
                UnitRateTotalRow = unitRateTotalRow;
                ImportWarning = importWarning;
                ImportNote = importNote;
            }

            public string LineId { get; }
            public string GroupKey { get; }
            public WorksheetRole UnitRateRole { get; }
            public int UnitRateTotalRow { get; }
            public string ImportWarning { get; }
            public string ImportNote { get; }
        }
    }
}
