using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookEstimateLine
    {
        internal WorkbookEstimateLine(
            string lineId,
            string groupKey,
            int targetRow,
            string description,
            string unit,
            decimal quantity,
            decimal acceptedQuantity,
            string quantitySourceReference,
            string outputReference,
            string normKey,
            string variantCode,
            string importWarning,
            string importNote,
            WorksheetRole unitRateRole,
            int unitRateTotalRow,
            IEnumerable<UnitRateResourceBinding> bindings)
        {
            LineId = lineId;
            GroupKey = groupKey;
            TargetRow = targetRow;
            Description = description;
            Unit = unit;
            Quantity = quantity;
            AcceptedQuantity = acceptedQuantity;
            QuantitySourceReference = quantitySourceReference;
            OutputReference = outputReference;
            NormKey = normKey;
            VariantCode = variantCode;
            ImportWarning = importWarning;
            ImportNote = importNote;
            UnitRateRole = unitRateRole;
            UnitRateTotalRow = unitRateTotalRow;
            Bindings = new ReadOnlyCollection<UnitRateResourceBinding>(bindings.ToList());
        }

        public string LineId { get; }
        public string GroupKey { get; }
        public int TargetRow { get; }
        public string Description { get; }
        public string Unit { get; }
        public decimal Quantity { get; }
        public decimal AcceptedQuantity { get; }
        public string QuantitySourceReference { get; }
        public string OutputReference { get; }
        public string NormKey { get; }
        public string VariantCode { get; }
        public string ImportWarning { get; }
        public string ImportNote { get; }
        public WorksheetRole UnitRateRole { get; }
        public int UnitRateTotalRow { get; }
        public IReadOnlyList<UnitRateResourceBinding> Bindings { get; }
        public bool IsActive => Quantity != 0m || AcceptedQuantity != 0m;
    }

    public sealed class WorkbookEstimatePlan
    {
        internal WorkbookEstimatePlan(
            string worksheetCodeName,
            IEnumerable<WorkbookEstimateLine> lines,
            int adjustmentRow,
            int grandTotalRow)
        {
            WorksheetCodeName = worksheetCodeName;
            Lines = new ReadOnlyCollection<WorkbookEstimateLine>(
                lines.OrderBy(line => line.TargetRow).ToList());
            AdjustmentRow = adjustmentRow;
            GrandTotalRow = grandTotalRow;
        }

        public string WorksheetCodeName { get; }
        public IReadOnlyList<WorkbookEstimateLine> Lines { get; }
        public int AdjustmentRow { get; }
        public int GrandTotalRow { get; }
    }

    public sealed class WorkbookEstimatePreviewLine
    {
        internal WorkbookEstimatePreviewLine(
            WorkbookEstimateLine line,
            UnitRateCalculationResult unitRate,
            UnitRateCalculationResult baselineUnitRate,
            string error,
            bool blocking)
        {
            Line = line;
            UnitRate = unitRate;
            BaselineUnitRate = baselineUnitRate;
            Error = error ?? string.Empty;
            IsBlocking = blocking;
        }

        public WorkbookEstimateLine Line { get; }
        public UnitRateCalculationResult UnitRate { get; }
        public UnitRateCalculationResult BaselineUnitRate { get; }
        public string Error { get; }
        public bool IsBlocking { get; }
    }

    public sealed class WorkbookEstimatePreview
    {
        internal WorkbookEstimatePreview(
            WorkbookEstimatePlan plan,
            IEnumerable<string> conditions,
            IEnumerable<WorkbookEstimatePreviewLine> lines,
            EstimateAppendixCalculationResult baselineResult,
            EstimateAppendixCalculationResult result)
        {
            Plan = plan;
            Conditions = new ReadOnlyCollection<string>(conditions.ToList());
            Lines = new ReadOnlyCollection<WorkbookEstimatePreviewLine>(lines.ToList());
            BaselineResult = baselineResult;
            Result = result;
        }

        public WorkbookEstimatePlan Plan { get; }
        public IReadOnlyList<string> Conditions { get; }
        public IReadOnlyList<WorkbookEstimatePreviewLine> Lines { get; }
        public EstimateAppendixCalculationResult BaselineResult { get; }
        public EstimateAppendixCalculationResult Result { get; }
        public bool IsValid => Lines.All(line => !line.IsBlocking) && Result != null;
    }

    public static class WorkbookEstimateAppendixService
    {
        private static readonly Regex NormCodePattern = new Regex(
            @"(?<code>\d{3}\.\d{4})(?:\.(?<variant>\d+))?",
            RegexOptions.CultureInvariant);

        public static WorkbookEstimatePlan ImportLegacyPlan(
            Excel.Workbook workbook,
            WorkbookUnitRateContext context)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            Excel.Worksheet target = null;
            Excel.Worksheet land = null;
            Excel.Worksheet water = null;
            Excel.Range used = null;
            Excel.Range targetData = null;
            Excel.Range targetRates = null;
            try
            {
                target = WorksheetRoleService.ResolveWorksheetRequired(
                    workbook,
                    WorksheetRole.EstimateAppendix);
                land = WorksheetRoleService.ResolveWorksheetRequired(workbook, WorksheetRole.UnitRateLand);
                water = WorksheetRoleService.ResolveWorksheetRequired(workbook, WorksheetRole.UnitRateWater);
                used = target.UsedRange;
                int firstRow = used.Row;
                int lastRow = firstRow + used.Rows.Count - 1;
                targetData = target.Range["B" + firstRow.ToString(CultureInfo.InvariantCulture),
                    "E" + lastRow.ToString(CultureInfo.InvariantCulture)];
                targetRates = target.Range["F" + firstRow.ToString(CultureInfo.InvariantCulture),
                    "H" + lastRow.ToString(CultureInfo.InvariantCulture)];
                object dataValues = targetData.Value2;
                object dataFormulas = targetData.Formula;
                object rateFormulas = targetRates.Formula;
                var lines = new List<WorkbookEstimateLine>();
                var occurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                for (int row = firstRow; row <= lastRow; row++)
                {
                    int offset = row - firstRow + 1;
                    string formulaF = Convert.ToString(ArrayValue(rateFormulas, offset, 1), CultureInfo.InvariantCulture);
                    string formulaG = Convert.ToString(ArrayValue(rateFormulas, offset, 2), CultureInfo.InvariantCulture);
                    string formulaH = Convert.ToString(ArrayValue(rateFormulas, offset, 3), CultureInfo.InvariantCulture);
                    WorksheetRole sourceRole;
                    Excel.Worksheet sourceSheet;
                    int sourceRow;
                    if (TryMatchRateTriplet(formulaF, formulaG, formulaH, land.Name, out sourceRow))
                    {
                        sourceRole = WorksheetRole.UnitRateLand;
                        sourceSheet = land;
                    }
                    else if (TryMatchRateTriplet(formulaF, formulaG, formulaH, water.Name, out sourceRow))
                    {
                        sourceRole = WorksheetRole.UnitRateWater;
                        sourceSheet = water;
                    }
                    else
                    {
                        continue;
                    }

                    NormIdentity identity = ReadNormIdentity(
                        sourceSheet,
                        sourceRow,
                        context.NormCatalog);
                    string description = Convert.ToString(
                        ArrayValue(dataValues, offset, 1),
                        CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
                    string variantCode = identity.VariantCode;
                    string importNote = string.Empty;
                    if (string.Equals(identity.NormKey, "NORM-030.0300", StringComparison.Ordinal) &&
                        Normalize(description).Contains("NUOC SAU 3M 12M") &&
                        !string.Equals(variantCode, "water-3-12", StringComparison.Ordinal))
                    {
                        variantCode = "water-3-12";
                        importNote = "Mo ta phu luc va hao phi legacy la do sau nuoc tren 3m den 12m; " +
                            "da anh xa sang water-3-12 thay cho hau to cu.";
                    }
                    string occurrenceKey = identity.NormKey + "|" + variantCode;
                    int occurrence;
                    occurrences.TryGetValue(occurrenceKey, out occurrence);
                    occurrence++;
                    occurrences[occurrenceKey] = occurrence;
                    string groupKey = sourceRole == WorksheetRole.UnitRateLand ? "LAND" : "WATER";
                    string lineId = groupKey + "|" + occurrenceKey + "|" +
                        occurrence.ToString(CultureInfo.InvariantCulture);
                    decimal quantity = ToDecimal(ArrayValue(dataValues, offset, 3), "D" + row);
                    decimal accepted = ToDecimal(ArrayValue(dataValues, offset, 4), "E" + row);
                    string quantityFormula = Convert.ToString(
                        ArrayValue(dataFormulas, offset, 3),
                        CultureInfo.InvariantCulture);
                    string quantitySource = quantityFormula != null && quantityFormula.StartsWith("=", StringComparison.Ordinal)
                        ? quantityFormula.Substring(1)
                        : target.CodeName + "!D" + row.ToString(CultureInfo.InvariantCulture);
                    IReadOnlyList<UnitRateResourceBinding> bindings = DetectLegacyBindings(
                        sourceSheet,
                        identity.HeaderRow,
                        sourceRow,
                        context.NormCatalog.FindRequired(identity.NormKey),
                        context.PriceProfile);
                    lines.Add(new WorkbookEstimateLine(
                        lineId,
                        groupKey,
                        row,
                        description,
                        Convert.ToString(ArrayValue(dataValues, offset, 2), CultureInfo.CurrentCulture)?.Trim() ?? string.Empty,
                        quantity,
                        accepted,
                        quantitySource,
                        target.CodeName + "!F" + row.ToString(CultureInfo.InvariantCulture) + ":N" +
                            row.ToString(CultureInfo.InvariantCulture),
                        identity.NormKey,
                        variantCode,
                        identity.Warning,
                        importNote,
                        sourceRole,
                        sourceRow,
                        bindings));
                }

                if (lines.Count == 0)
                {
                    WorkbookEstimatePlan restoredPlan;
                    if (WorkbookResultAuditService.TryRestoreEstimatePlan(
                        workbook,
                        target,
                        context,
                        out restoredPlan))
                    {
                        return restoredPlan;
                    }
                    throw new InvalidOperationException(
                        "Khong tim thay dong phu luc lien ket voi bang don gia va workbook chua co audit de khoi phuc.");
                }
                int maxLine = lines.Max(line => line.TargetRow);
                int adjustmentRow = FindAdjustmentRow(target, maxLine + 1, lastRow);
                int grandTotalRow = FindGrandTotalRow(target, maxLine + 1, lastRow);
                return new WorkbookEstimatePlan(target.CodeName, lines, adjustmentRow, grandTotalRow);
            }
            finally
            {
                Release(targetRates);
                Release(targetData);
                Release(used);
                Release(water);
                Release(land);
                Release(target);
            }
        }

        public static WorkbookEstimatePreview Preview(
            WorkbookUnitRateContext context,
            WorkbookEstimatePlan plan,
            IEnumerable<string> waterConditions)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (plan == null)
                throw new ArgumentNullException(nameof(plan));
            string[] conditions = (waterConditions ?? Enumerable.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
            var previewLines = new List<WorkbookEstimatePreviewLine>();
            var coreLines = new List<EstimateAppendixLineRequest>();
            var baselineCoreLines = new List<EstimateAppendixLineRequest>();
            bool hasBlockingIssue = false;
            foreach (WorkbookEstimateLine line in plan.Lines)
            {
                UnitRateCalculationResult rate = null;
                UnitRateCalculationResult baselineRate = null;
                string error = string.Empty;
                bool blocking = false;
                if (!string.IsNullOrEmpty(line.ImportWarning))
                {
                    error = line.ImportWarning;
                    blocking = line.IsActive;
                }
                else try
                {
                    rate = WorkbookUnitRateService.Calculate(
                        context,
                        line.NormKey,
                        line.VariantCode,
                        line.GroupKey == "WATER" ? conditions : Array.Empty<string>(),
                        line.Bindings.ToArray());
                    baselineRate = line.GroupKey == "WATER" && conditions.Length > 0
                        ? WorkbookUnitRateService.Calculate(
                            context,
                            line.NormKey,
                            line.VariantCode,
                            Array.Empty<string>(),
                            line.Bindings.ToArray())
                        : rate;
                }
                catch (UnitRateValidationException ex)
                {
                    error = ex.Message;
                    blocking = line.IsActive;
                }
                catch (ArgumentException ex)
                {
                    error = ex.Message;
                    blocking = line.IsActive;
                }
                catch (InvalidOperationException ex)
                {
                    error = ex.Message;
                    blocking = line.IsActive;
                }
                if (!blocking && string.IsNullOrEmpty(error) && !string.IsNullOrEmpty(line.ImportNote))
                    error = line.ImportNote;
                hasBlockingIssue |= blocking;
                previewLines.Add(new WorkbookEstimatePreviewLine(
                    line,
                    rate,
                    baselineRate,
                    error,
                    blocking));
                coreLines.Add(new EstimateAppendixLineRequest(
                    line.LineId,
                    line.GroupKey,
                    line.Description,
                    line.Unit,
                    line.Quantity,
                    line.AcceptedQuantity,
                    rate,
                    line.QuantitySourceReference,
                    line.OutputReference));
                baselineCoreLines.Add(new EstimateAppendixLineRequest(
                    line.LineId,
                    line.GroupKey,
                    line.Description,
                    line.Unit,
                    line.Quantity,
                    line.AcceptedQuantity,
                    baselineRate,
                    line.QuantitySourceReference,
                    line.OutputReference));
            }

            EstimateAppendixCalculationResult result = null;
            EstimateAppendixCalculationResult baselineResult = null;
            if (!hasBlockingIssue)
            {
                baselineResult = EstimateAppendixCalculator.Calculate(
                    new EstimateAppendixCalculationRequest(
                        context.PackageId,
                        context.PackageVersion,
                        context.PackageChecksum,
                        baselineCoreLines));
                result = EstimateAppendixCalculator.Calculate(
                    new EstimateAppendixCalculationRequest(
                        context.PackageId,
                        context.PackageVersion,
                        context.PackageChecksum,
                        coreLines));
            }
            return new WorkbookEstimatePreview(
                plan,
                conditions,
                previewLines,
                baselineResult,
                result);
        }

        private static bool TryMatchRateTriplet(
            string formulaF,
            string formulaG,
            string formulaH,
            string sheetName,
            out int row)
        {
            int rowF;
            int rowG;
            int rowH;
            bool valid = TryParseReference(formulaF, sheetName, "F", out rowF) &&
                TryParseReference(formulaG, sheetName, "G", out rowG) &&
                TryParseReference(formulaH, sheetName, "H", out rowH) &&
                rowF == rowG && rowF == rowH;
            row = valid ? rowF : 0;
            return valid;
        }

        private static bool TryParseReference(
            string formula,
            string sheetName,
            string column,
            out int row)
        {
            row = 0;
            if (string.IsNullOrWhiteSpace(formula))
                return false;
            string escapedSheet = Regex.Escape(sheetName.Replace("'", "''"));
            Match match = Regex.Match(
                formula.Trim(),
                @"^[=+]*'?" + escapedSheet + @"'?!\$?" + column + @"\$?(?<row>\d+)$",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success && int.TryParse(
                match.Groups["row"].Value,
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out row);
        }

        private static NormIdentity ReadNormIdentity(
            Excel.Worksheet worksheet,
            int totalRow,
            NormCatalog catalog)
        {
            Excel.Range range = null;
            try
            {
                int startRow = Math.Max(1, totalRow - 40);
                range = worksheet.Range[
                    "A" + startRow.ToString(CultureInfo.InvariantCulture),
                    "A" + totalRow.ToString(CultureInfo.InvariantCulture)];
                object values = range.Value2;
                int count = totalRow - startRow + 1;
                for (int offset = count; offset >= 1; offset--)
                {
                    string text = Convert.ToString(ArrayValue(values, offset, 1), CultureInfo.CurrentCulture);
                    Match match = NormCodePattern.Match(text ?? string.Empty);
                    if (!match.Success)
                        continue;
                    string normKey = "NORM-" + match.Groups["code"].Value;
                    NormDefinition definition = catalog.FindRequired(normKey);
                    int variantIndex = 1;
                    if (match.Groups["variant"].Success && !int.TryParse(
                        match.Groups["variant"].Value,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out variantIndex))
                    {
                        throw new FormatException("Hau to bien the dinh muc khong hop le: " + text + ".");
                    }
                    string variantCode = ResolveLegacyVariant(normKey, definition, variantIndex);
                    if (variantCode == null)
                    {
                        return new NormIdentity(
                            normKey,
                            string.Empty,
                            startRow + offset - 1,
                            "Bien the legacy nam ngoai catalog phap ly: " + text +
                                ". Phai chon lai bien the khi dong co khoi luong.");
                    }
                    return new NormIdentity(
                        normKey,
                        variantCode,
                        startRow + offset - 1,
                        string.Empty);
                }
                throw new InvalidOperationException(
                    "Khong tim thay so hieu dinh muc truoc dong tong " + totalRow +
                    " cua sheet " + worksheet.Name + ".");
            }
            finally
            {
                Release(range);
            }
        }

        private static IReadOnlyList<UnitRateResourceBinding> DetectLegacyBindings(
            Excel.Worksheet worksheet,
            int headerRow,
            int totalRow,
            NormDefinition definition,
            PriceProfile profile)
        {
            string[] logicalCodes = definition.Rates
                .Select(rate => rate.ResourceCode)
                .Where(UnitRateCalculator.IsLogicalResource)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (logicalCodes.Length == 0)
                return Array.Empty<UnitRateResourceBinding>();

            Excel.Range namesRange = null;
            try
            {
                namesRange = worksheet.Range[
                    "B" + headerRow.ToString(CultureInfo.InvariantCulture),
                    "B" + totalRow.ToString(CultureInfo.InvariantCulture)];
                object names = namesRange.Value2;
                var machinePrices = profile.Entries
                    .Where(entry => entry.Kind == PriceResourceKind.MachineShift)
                    .ToArray();
                var fixedMachinePriceCodes = new HashSet<string>(
                    definition.Rates
                        .Select(rate => rate.ResourceCode)
                        .Where(code => !UnitRateCalculator.IsLogicalResource(code))
                        .Select(code => "MACHINE-" + code),
                    StringComparer.OrdinalIgnoreCase);
                var bindings = new List<UnitRateResourceBinding>();
                foreach (string logicalCode in logicalCodes)
                {
                    UnitRateResourceBinding binding = null;
                    int count = totalRow - headerRow + 1;
                    for (int offset = 1; offset <= count && binding == null; offset++)
                    {
                        string legacyName = Convert.ToString(
                            ArrayValue(names, offset, 1),
                            CultureInfo.CurrentCulture)?.Trim();
                        if (string.IsNullOrEmpty(legacyName))
                            continue;
                        string normalizedLegacy = Normalize(legacyName);
                        PriceProfileEntry price = machinePrices
                            .Where(candidate => !fixedMachinePriceCodes.Contains(candidate.Code))
                            .Select(candidate => new
                            {
                                Price = candidate,
                                Normalized = Normalize(candidate.DisplayName)
                            })
                            .Where(candidate =>
                                normalizedLegacy == candidate.Normalized ||
                                normalizedLegacy.Contains(candidate.Normalized) ||
                                candidate.Normalized.Contains(normalizedLegacy))
                            .OrderBy(candidate => normalizedLegacy == candidate.Normalized ? 0 : 1)
                            .ThenBy(candidate => Math.Abs(normalizedLegacy.Length - candidate.Normalized.Length))
                            .ThenBy(candidate => candidate.Price.Code, StringComparer.Ordinal)
                            .Select(candidate => candidate.Price)
                            .FirstOrDefault();
                        if (price != null)
                        {
                            int sourceRow = headerRow + offset - 1;
                            binding = new UnitRateResourceBinding(
                                logicalCode,
                                price.Code,
                                "Nhap tu bang don gia legacy " + worksheet.CodeName + "!B" +
                                    sourceRow.ToString(CultureInfo.InvariantCulture));
                        }
                    }
                    if (binding != null)
                        bindings.Add(binding);
                }
                return bindings;
            }
            finally
            {
                Release(namesRange);
            }
        }

        private static string ResolveLegacyVariant(
            string normKey,
            NormDefinition definition,
            int legacyVariantIndex)
        {
            if (string.Equals(normKey, "NORM-020.0500", StringComparison.Ordinal))
            {
                switch (legacyVariantIndex)
                {
                    case 1: return "depth-0.5-or-1";
                    case 2: return "depth-3";
                    case 4: return "depth-5";
                    case 5: return "depth-10";
                    default: return null;
                }
            }
            if (legacyVariantIndex < 1 || legacyVariantIndex > definition.Variants.Count)
                return null;
            return definition.Variants[legacyVariantIndex - 1];
        }

        private static int FindAdjustmentRow(Excel.Worksheet worksheet, int startRow, int lastRow)
        {
            for (int row = startRow; row <= lastRow; row++)
            {
                Excel.Range descriptionCell = null;
                Excel.Range laborCell = null;
                try
                {
                    descriptionCell = worksheet.Cells[row, 2];
                    laborCell = worksheet.Cells[row, 10];
                    string description = Convert.ToString(descriptionCell.Value2, CultureInfo.CurrentCulture);
                    string formula = Convert.ToString(laborCell.Formula, CultureInfo.InvariantCulture);
                    if (Normalize(description).Contains("HE SO") &&
                        !string.IsNullOrWhiteSpace(formula) &&
                        formula.StartsWith("=", StringComparison.Ordinal))
                    {
                        return row;
                    }
                }
                finally
                {
                    Release(laborCell);
                    Release(descriptionCell);
                }
            }
            return 0;
        }

        private static int FindGrandTotalRow(Excel.Worksheet worksheet, int startRow, int lastRow)
        {
            for (int row = startRow; row <= lastRow; row++)
            {
                Excel.Range cell = null;
                try
                {
                    cell = worksheet.Cells[row, 2];
                    if (Normalize(Convert.ToString(cell.Value2, CultureInfo.CurrentCulture)) == "TONG CONG")
                        return row;
                }
                finally
                {
                    Release(cell);
                }
            }
            return 0;
        }

        private static decimal ToDecimal(object value, string address)
        {
            if (value == null || value is string && string.IsNullOrWhiteSpace((string)value))
                return 0m;
            try
            {
                decimal result = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                if (result < 0m)
                    throw new InvalidOperationException("Khoi luong am tai " + address + ".");
                return result;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Khoi luong khong phai so tai " + address + ".", ex);
            }
        }

        private static object ArrayValue(object values, int row, int column)
        {
            Array array = values as Array;
            if (array != null)
                return array.GetValue(row, column);
            return row == 1 && column == 1 ? values : null;
        }

        private static string Normalize(string value)
        {
            string normalized = (value ?? string.Empty).Normalize(NormalizationForm.FormD);
            var output = new StringBuilder(normalized.Length);
            bool previousSpace = true;
            foreach (char character in normalized)
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(character);
                if (category == UnicodeCategory.NonSpacingMark)
                    continue;
                if (char.IsLetterOrDigit(character))
                {
                    output.Append(char.ToUpperInvariant(character));
                    previousSpace = false;
                }
                else if (!previousSpace)
                {
                    output.Append(' ');
                    previousSpace = true;
                }
            }
            return output.ToString().Trim();
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private sealed class NormIdentity
        {
            internal NormIdentity(string normKey, string variantCode, int headerRow, string warning)
            {
                NormKey = normKey;
                VariantCode = variantCode;
                HeaderRow = headerRow;
                Warning = warning;
            }

            internal string NormKey { get; }
            internal string VariantCode { get; }
            internal int HeaderRow { get; }
            internal string Warning { get; }
        }
    }
}
