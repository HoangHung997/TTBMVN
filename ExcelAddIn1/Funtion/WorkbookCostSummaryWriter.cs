using ExcelAddIn1.Core;
using System;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookCostSummaryWriteResult
    {
        internal WorkbookCostSummaryWriteResult(
            string worksheetName,
            string writtenAddress,
            CostSummaryCalculationResult calculation)
        {
            WorksheetName = worksheetName;
            WrittenAddress = writtenAddress;
            Calculation = calculation;
        }

        public string WorksheetName { get; }
        public string WrittenAddress { get; }
        public CostSummaryCalculationResult Calculation { get; }
    }

    public static class WorkbookCostSummaryWriter
    {
        public static WorkbookCostSummaryWriteResult Apply(
            Excel.Workbook workbook,
            WorkbookCostSummaryPreview preview)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (preview == null)
                throw new ArgumentNullException(nameof(preview));

            Excel.Worksheet summary = null;
            Excel.Worksheet estimate = null;
            Excel.Range calculationRange = null;
            Excel.Range wordsRange = null;
            try
            {
                summary = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.CostSummary);
                estimate = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.EstimateAppendix);
                if (!string.Equals(summary.CodeName, preview.Context.SummarySheetCodeName, StringComparison.Ordinal) ||
                    !string.Equals(estimate.CodeName, preview.Context.EstimateSheetCodeName, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Vai trò sheet đã thay đổi từ lúc xem trước. Hãy tải lại bảng tổng hợp.");
                }

                calculationRange = summary.Range["D10:E27"];
                wordsRange = summary.Range["A28"];
                object[,] formulas = CloneMatrix(calculationRange.Formula, 18, 2);
                Fill(formulas, estimate.Name, preview);

                using (new ExcelWriteContext(workbook.Application))
                using (var transaction = new ExcelBatchWriteTransaction())
                {
                    transaction.WriteFormula(calculationRange, formulas);
                    transaction.WriteValue2(
                        wordsRange,
                        "Bằng chữ: " + VietnameseMoneyWords.ToWords(preview.Result.RoundedAfterTaxVnd));
                    calculationRange.Calculate();
                    Verify(summary, preview.Result.RoundedAfterTaxVnd);
                    WorkbookResultAuditService.SaveScope(
                        workbook,
                        WorkbookResultAuditService.CostSummaryScopeId,
                        WorkbookResultAuditService.BuildCostSummaryEntries(
                            workbook,
                            preview,
                            summary.CodeName));
                    transaction.Commit();
                }

                return new WorkbookCostSummaryWriteResult(
                    summary.Name,
                    "D10:E27;A28",
                    preview.Result);
            }
            finally
            {
                Release(wordsRange);
                Release(calculationRange);
                Release(estimate);
                Release(summary);
            }
        }

        internal static IReadOnlyDictionary<string, string> BuildExpectedFormulas(
            string estimateSheetName,
            WorkbookCostSummaryPreview preview)
        {
            if (preview == null)
                throw new ArgumentNullException(nameof(preview));
            var output = new object[18, 2];
            Fill(output, estimateSheetName, preview);
            var formulas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int row = 10; row <= 27; row++)
            {
                for (int column = 4; column <= 5; column++)
                {
                    string formula = output[row - 10, column - 4] as string;
                    if (formula != null && formula.StartsWith("=", StringComparison.Ordinal))
                        formulas.Add(
                            ExcelColumnAddress.ToLetters(column) + row.ToString(CultureInfo.InvariantCulture),
                            formula);
                }
            }
            return formulas;
        }

        private static void Fill(
            object[,] values,
            string estimateSheetName,
            WorkbookCostSummaryPreview preview)
        {
            CostSummaryCalculationRequest request = preview.Request;
            CostSummaryCalculationResult result = preview.Result;
            string sheet = "'" + (estimateSheetName ?? string.Empty).Replace("'", "''") + "'";

            Set(values, 10, 5, "=ROUND(" + sheet + "!I27,0)");
            Set(values, 11, 5, "=ROUND(" + sheet + "!J27,0)");
            Set(values, 12, 5, "=ROUND(" + sheet + "!K27,0)");
            Set(values, 13, 5, "=SUM(E10:E12)");
            Set(values, 14, 5, "=ROUND(" +
                preview.Context.Catalog.CommonRatePercent.ToString(CultureInfo.InvariantCulture) + "%*E11,0)");
            Set(values, 15, 4, Number(request.PreTaxIncomeRatePercent));
            Set(values, 15, 5, request.Template == CostSummaryTemplate.OtherFunding
                ? "=ROUND(D15%*(E13+E14),0)"
                : "=0");
            Set(values, 16, 5, request.Template == CostSummaryTemplate.OtherFunding
                ? "=SUM(E13:E15)"
                : "=SUM(E13:E14)");

            foreach (string code in new[] { "K1", "K2", "K3", "K4", "K5", "K6" })
            {
                int row = 17 + (code[1] - '0');
                CostComponentCalculationResult component = Find(result, code);
                Set(values, row, 4, component == null || !component.CalculatedRatePercent.HasValue
                    ? 0d
                    : Number(component.CalculatedRatePercent.Value));
                Set(values, row, 5, FormulaForComponent(row, component, request, preview.Context.Catalog));
            }
            Set(values, 17, 5, "=SUM(E18:E23)");
            Set(values, 24, 5, "=E16+E17");
            Set(values, 25, 4, request.Template == CostSummaryTemplate.OtherFunding
                ? request.VatRatePercent.ToString("0.####", CultureInfo.InvariantCulture) + "%*(Q-K3-K4)"
                : "Không áp dụng");
            Set(values, 25, 5, request.Template == CostSummaryTemplate.OtherFunding
                ? "=ROUND(" + request.VatRatePercent.ToString(CultureInfo.InvariantCulture) + "%*(E24-E20-E21),0)"
                : "=0");
            Set(values, 26, 5, "=E24+E25");
            Set(values, 27, 5, "=ROUND(E26,-3)");
        }

        private static object FormulaForComponent(
            int row,
            CostComponentCalculationResult component,
            CostSummaryCalculationRequest request,
            CostRuleCatalog catalog)
        {
            if (component == null)
                return "=0";
            if (component.IsOverridden)
                return component.AppliedAmountVnd;
            switch (component.ComponentCode)
            {
                case "K1":
                    return request.AreaHa <= catalog.MinimumAreaHa
                        ? "=MAX(ROUND(D18%*E16,0)," + catalog.MinimumSurveyVnd + ")"
                        : "=ROUND(D18%*E16,0)";
                case "K2": return "=ROUND(D19%*E13,0)";
                case "K3": return "=MAX(" + catalog.K3MinimumVnd + ",MIN(" +
                    catalog.K3MaximumVnd + ",ROUND(D20%*E16,0)))";
                case "K4": return "=ROUND(D21%*E16,0)";
                case "K5": return "=ROUND(D22%*E16,0)";
                case "K6":
                    return request.AreaHa <= catalog.MinimumAreaHa
                        ? "=MAX(ROUND(D23%*E16,0)," + catalog.MinimumDisposalVnd + ")"
                        : "=ROUND(D23%*E16,0)";
                default: throw new InvalidOperationException("Không hỗ trợ " + component.ComponentCode + ".");
            }
        }

        private static CostComponentCalculationResult Find(
            CostSummaryCalculationResult result,
            string code)
        {
            foreach (CostComponentCalculationResult component in result.Components)
            {
                if (string.Equals(component.ComponentCode, code, StringComparison.Ordinal))
                    return component;
            }
            return null;
        }

        private static void Verify(Excel.Worksheet worksheet, long expected)
        {
            Excel.Range range = null;
            try
            {
                range = worksheet.Range["E27"];
                long actual = Convert.ToInt64(
                    Convert.ToDecimal(range.Value2 ?? 0d, CultureInfo.InvariantCulture));
                if (actual != expected)
                {
                    throw new InvalidOperationException(
                        "Kiểm tra tổng kinh phí thất bại: expected=" + expected + ", actual=" + actual + ".");
                }
            }
            finally
            {
                Release(range);
            }
        }

        private static object[,] CloneMatrix(object source, int rows, int columns)
        {
            var output = new object[rows, columns];
            Array values = source as Array;
            if (values == null)
                return output;
            int rowLower = values.GetLowerBound(0);
            int columnLower = values.GetLowerBound(1);
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                    output[row, column] = values.GetValue(rowLower + row, columnLower + column);
            }
            return output;
        }

        private static void Set(object[,] values, int row, int column, object value)
        {
            values[row - 10, column - 4] = value;
        }

        private static double Number(decimal value)
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
