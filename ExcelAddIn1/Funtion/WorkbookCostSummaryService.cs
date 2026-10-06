using ExcelAddIn1.Core;
using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookCostSummaryContext
    {
        internal WorkbookCostSummaryContext(
            string packageId,
            string packageVersion,
            string packageChecksum,
            CostRuleCatalog catalog,
            string summarySheetCodeName,
            string estimateSheetCodeName,
            CostSummaryCalculationRequest defaultRequest,
            string importNote)
        {
            PackageId = packageId;
            PackageVersion = packageVersion;
            PackageChecksum = packageChecksum;
            Catalog = catalog;
            SummarySheetCodeName = summarySheetCodeName;
            EstimateSheetCodeName = estimateSheetCodeName;
            DefaultRequest = defaultRequest;
            ImportNote = importNote ?? string.Empty;
        }

        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public CostRuleCatalog Catalog { get; }
        public string SummarySheetCodeName { get; }
        public string EstimateSheetCodeName { get; }
        public CostSummaryCalculationRequest DefaultRequest { get; }
        public string ImportNote { get; }
    }

    public sealed class WorkbookCostSummaryPreview
    {
        internal WorkbookCostSummaryPreview(
            WorkbookCostSummaryContext context,
            CostSummaryCalculationRequest request,
            CostSummaryCalculationResult result)
        {
            Context = context;
            Request = request;
            Result = result;
        }

        public WorkbookCostSummaryContext Context { get; }
        public CostSummaryCalculationRequest Request { get; }
        public CostSummaryCalculationResult Result { get; }
    }

    public static class WorkbookCostSummaryService
    {
        public static WorkbookCostSummaryContext Load(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            WorkbookCostRuleContext ruleContext = WorkbookCostRuleService.LoadPinnedCatalog(workbook);
            Excel.Worksheet summary = null;
            Excel.Worksheet estimate = null;
            try
            {
                summary = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.CostSummary);
                estimate = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.EstimateAppendix);

                decimal material = ReadDecimal(estimate, "I27");
                decimal labor = ReadDecimal(estimate, "J27");
                decimal machine = ReadDecimal(estimate, "K27");
                decimal area = Math.Max(0m, ReadDecimal(estimate, "D13") + ReadDecimal(estimate, "D20"));
                if (area == 0m)
                    area = 5m;

                long direct = CostRuleCalculator.CalculateDirect(
                    ruleContext.Catalog, material, labor, machine).DirectVnd;
                string terrain = InferTerrain(ruleContext.Catalog, ReadDecimal(summary, "D18"));
                CostProjectKind project = InferProjectKind(
                    ruleContext.Catalog,
                    direct,
                    ReadDecimal(summary, "D19"));
                string notes = ReadText(summary, "G15:K15") + " " + ReadText(summary, "G22:K22");
                CostConstructionKind construction = InferConstructionKind(notes);
                decimal disposalWeight = InferDisposalWeight(
                    ruleContext.Catalog,
                    ReadDecimal(summary, "D23"));
                decimal preTaxIncomeRate = ReadDecimal(summary, "D15");
                if (preTaxIncomeRate == 0m)
                    preTaxIncomeRate = 5.5m;
                decimal vatRate = InferPercent(ReadFormula(summary, "E25"), 8m);

                var request = new CostSummaryCalculationRequest(
                    CostSummaryTemplate.OtherFunding,
                    material,
                    labor,
                    machine,
                    terrain,
                    area,
                    project,
                    construction,
                    disposalWeight,
                    preTaxIncomeRate,
                    vatRate,
                    CostComponentSelection.All);
                string note =
                    "Đã nhập cấu hình cũ từ THKP-TC. Diện tích được lấy từ khối lượng dò tìm cạn và nước; " +
                    "khối lượng hủy nổ được nhập theo nhóm dưới/trên 1.000 kg.";
                return new WorkbookCostSummaryContext(
                    ruleContext.PackageId,
                    ruleContext.PackageVersion,
                    ruleContext.PackageChecksum,
                    ruleContext.Catalog,
                    summary.CodeName,
                    estimate.CodeName,
                    request,
                    note);
            }
            finally
            {
                Release(estimate);
                Release(summary);
            }
        }

        public static WorkbookCostSummaryPreview Preview(
            WorkbookCostSummaryContext context,
            CostSummaryCalculationRequest request)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (request == null)
                throw new ArgumentNullException(nameof(request));
            return new WorkbookCostSummaryPreview(
                context,
                request,
                CostSummaryCalculator.Calculate(context.Catalog, request));
        }

        private static string InferTerrain(CostRuleCatalog catalog, decimal legacyRate)
        {
            CostTerrainDefinition exact = catalog.TerrainDefinitions.FirstOrDefault(item =>
                item.K1Percent == legacyRate);
            return exact == null ? "plain-open" : exact.Terrain;
        }

        private static CostProjectKind InferProjectKind(
            CostRuleCatalog catalog,
            decimal direct,
            decimal legacyRate)
        {
            return new[] { CostProjectKind.Linear, CostProjectKind.Other }
                .OrderBy(kind => Math.Abs(
                    CostRuleCalculator.CalculateK2(catalog, kind, direct).RatePercent - legacyRate))
                .First();
        }

        private static CostConstructionKind InferConstructionKind(string value)
        {
            string text = RemoveDiacritics(value).ToLowerInvariant();
            if (text.Contains("nong nghiep") || text.Contains("moi truong"))
                return CostConstructionKind.AgricultureAndEnvironment;
            if (text.Contains("giao thong"))
                return CostConstructionKind.Transport;
            if (text.Contains("cong nghiep"))
                return CostConstructionKind.Industrial;
            if (text.Contains("ha tang"))
                return CostConstructionKind.TechnicalInfrastructure;
            return CostConstructionKind.Civil;
        }

        private static decimal InferDisposalWeight(CostRuleCatalog catalog, decimal rate)
        {
            return Math.Abs(rate - catalog.K6Under1000Percent) <=
                Math.Abs(rate - catalog.K6Over1000Percent)
                ? 999m
                : 1001m;
        }

        private static decimal InferPercent(string formula, decimal fallback)
        {
            Match match = Regex.Match(formula ?? string.Empty, @"(?<rate>\d+(?:[\.,]\d+)?)\s*%");
            if (!match.Success)
                return fallback;
            string value = match.Groups["rate"].Value.Replace(',', '.');
            return decimal.TryParse(
                value,
                NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture,
                out decimal parsed)
                ? parsed
                : fallback;
        }

        private static decimal ReadDecimal(Excel.Worksheet worksheet, string address)
        {
            Excel.Range range = null;
            try
            {
                range = worksheet.Range[address];
                return Convert.ToDecimal(range.Value2 ?? 0d, CultureInfo.InvariantCulture);
            }
            finally
            {
                Release(range);
            }
        }

        private static string ReadFormula(Excel.Worksheet worksheet, string address)
        {
            Excel.Range range = null;
            try
            {
                range = worksheet.Range[address];
                return Convert.ToString(range.Formula, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            finally
            {
                Release(range);
            }
        }

        private static string ReadText(Excel.Worksheet worksheet, string address)
        {
            Excel.Range range = null;
            try
            {
                range = worksheet.Range[address];
                object raw = range.Value2;
                Array values = raw as Array;
                if (values == null)
                    return Convert.ToString(raw, CultureInfo.CurrentCulture) ?? string.Empty;
                var builder = new StringBuilder();
                foreach (object item in values)
                    builder.Append(' ').Append(Convert.ToString(item, CultureInfo.CurrentCulture));
                return builder.ToString();
            }
            finally
            {
                Release(range);
            }
        }

        private static string RemoveDiacritics(string value)
        {
            string normalized = (value ?? string.Empty).Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder(normalized.Length);
            foreach (char character in normalized)
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(character);
                if (category != UnicodeCategory.NonSpacingMark)
                    builder.Append(character == 'Đ' ? 'D' : character == 'đ' ? 'd' : character);
            }
            return builder.ToString().Normalize(NormalizationForm.FormC);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
