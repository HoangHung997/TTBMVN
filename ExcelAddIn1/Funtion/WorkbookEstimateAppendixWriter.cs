using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookEstimateWriteResult
    {
        internal WorkbookEstimateWriteResult(
            string worksheetCodeName,
            int writtenLineCount,
            int firstRow,
            int lastRow)
        {
            WorksheetCodeName = worksheetCodeName;
            WrittenLineCount = writtenLineCount;
            FirstRow = firstRow;
            LastRow = lastRow;
        }

        public string WorksheetCodeName { get; }
        public int WrittenLineCount { get; }
        public int FirstRow { get; }
        public int LastRow { get; }
    }

    public static class WorkbookEstimateAppendixWriter
    {
        private const decimal Tolerance = 0.0001m;

        public static WorkbookEstimateWriteResult Apply(
            Excel.Workbook workbook,
            WorkbookEstimatePreview preview)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (preview == null)
                throw new ArgumentNullException(nameof(preview));
            if (!preview.IsValid || preview.BaselineResult == null)
                throw new InvalidOperationException("Preview phu luc con loi, khong the ghi Excel.");
            if (preview.Plan.Lines.Count == 0)
                throw new InvalidOperationException("Phu luc khong co dong de ghi.");

            Excel.Worksheet worksheet = null;
            Excel.Range writeRange = null;
            try
            {
                worksheet = WorksheetRoleService.ResolveWorksheetRequired(
                    workbook,
                    WorksheetRole.EstimateAppendix);
                if (!string.Equals(
                    worksheet.CodeName,
                    preview.Plan.WorksheetCodeName,
                    StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Sheet phu luc da thay doi sau khi preview. Hay nap lai truoc khi ghi.");
                }

                int firstLineRow = preview.Plan.Lines.Min(line => line.TargetRow);
                int firstRow = Math.Max(1, firstLineRow - 1);
                int lastRow = Math.Max(
                    preview.Plan.Lines.Max(line => line.TargetRow),
                    Math.Max(preview.Plan.AdjustmentRow, preview.Plan.GrandTotalRow));
                writeRange = worksheet.Range[
                    "F" + firstRow.ToString(CultureInfo.InvariantCulture),
                    "N" + lastRow.ToString(CultureInfo.InvariantCulture)];
                object[,] output = CloneMatrix(writeRange.Formula, lastRow - firstRow + 1, 9);
                PopulateLineRows(output, firstRow, preview);
                PopulateGroupRows(output, firstRow, preview);
                PopulateAdjustmentRow(output, firstRow, preview);
                PopulateGrandTotalRow(output, firstRow, preview);

                using (new ExcelWriteContext(workbook.Application))
                using (var transaction = new ExcelBatchWriteTransaction())
                {
                    transaction.WriteFormula(writeRange, output);
                    worksheet.Calculate();
                    VerifyGrandTotal(worksheet, preview);
                    WorkbookResultAuditService.SaveScope(
                        workbook,
                        WorkbookResultAuditService.EstimateScopeId,
                        WorkbookResultAuditService.BuildEstimateEntries(
                            preview,
                            worksheet.CodeName,
                            firstRow,
                            lastRow));
                    transaction.Commit();
                }

                return new WorkbookEstimateWriteResult(
                    worksheet.CodeName,
                    preview.Plan.Lines.Count,
                    firstRow,
                    lastRow);
            }
            finally
            {
                Release(writeRange);
                Release(worksheet);
            }
        }

        internal static IReadOnlyDictionary<string, string> BuildExpectedFormulas(
            WorkbookEstimatePreview preview)
        {
            if (preview == null || !preview.IsValid)
                throw new ArgumentException("Preview phu luc khong hop le.", nameof(preview));
            int firstLineRow = preview.Plan.Lines.Min(line => line.TargetRow);
            int firstRow = Math.Max(1, firstLineRow - 1);
            int lastRow = Math.Max(
                preview.Plan.Lines.Max(line => line.TargetRow),
                Math.Max(preview.Plan.AdjustmentRow, preview.Plan.GrandTotalRow));
            var output = new object[lastRow - firstRow + 1, 9];
            PopulateLineRows(output, firstRow, preview);
            PopulateGroupRows(output, firstRow, preview);
            PopulateAdjustmentRow(output, firstRow, preview);
            PopulateGrandTotalRow(output, firstRow, preview);

            var formulas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int column = 6; column <= 14; column++)
                {
                    string formula = output[row - firstRow, column - 6] as string;
                    if (formula != null && formula.StartsWith("=", StringComparison.Ordinal))
                        formulas.Add(ColumnName(column) + row.ToString(CultureInfo.InvariantCulture), formula);
                }
            }
            return formulas;
        }

        private static void PopulateLineRows(
            object[,] output,
            int firstRow,
            WorkbookEstimatePreview preview)
        {
            var baselineLines = preview.BaselineResult.Lines.ToDictionary(
                line => line.Request.LineId,
                StringComparer.OrdinalIgnoreCase);
            foreach (WorkbookEstimateLine line in preview.Plan.Lines)
            {
                EstimateAppendixLineResult baseline = baselineLines[line.LineId];
                UnitRateCalculationResult rate = baseline.Request.UnitRate;
                Set(output, firstRow, line.TargetRow, 6, Number(rate?.MaterialAmountVnd ?? 0m));
                Set(output, firstRow, line.TargetRow, 7, Number(rate?.LaborAmountVnd ?? 0m));
                Set(output, firstRow, line.TargetRow, 8, Number(rate?.MachineAmountVnd ?? 0m));
                for (int column = 9; column <= 11; column++)
                {
                    Set(output, firstRow, line.TargetRow, column,
                        "=D" + line.TargetRow.ToString(CultureInfo.InvariantCulture) + "*" +
                        ColumnName(column - 3) + line.TargetRow.ToString(CultureInfo.InvariantCulture));
                }
                for (int column = 12; column <= 14; column++)
                {
                    Set(output, firstRow, line.TargetRow, column,
                        "=E" + line.TargetRow.ToString(CultureInfo.InvariantCulture) + "*" +
                        ColumnName(column - 6) + line.TargetRow.ToString(CultureInfo.InvariantCulture));
                }
            }
        }

        private static void PopulateGroupRows(
            object[,] output,
            int firstRow,
            WorkbookEstimatePreview preview)
        {
            foreach (IGrouping<string, WorkbookEstimateLine> group in preview.Plan.Lines
                .GroupBy(line => line.GroupKey, StringComparer.Ordinal))
            {
                int start = group.Min(line => line.TargetRow);
                int end = group.Max(line => line.TargetRow);
                int subtotalRow = start - 1;
                if (subtotalRow < firstRow)
                    throw new InvalidOperationException("Khong xac dinh duoc dong cong nhom " + group.Key + ".");
                int sumEnd = string.Equals(group.Key, "WATER", StringComparison.Ordinal) &&
                    preview.Plan.AdjustmentRow > end
                        ? preview.Plan.AdjustmentRow
                        : end;
                for (int column = 9; column <= 14; column++)
                {
                    string name = ColumnName(column);
                    Set(output, firstRow, subtotalRow, column,
                        "=SUM(" + name + start.ToString(CultureInfo.InvariantCulture) + ":" +
                        name + sumEnd.ToString(CultureInfo.InvariantCulture) + ")");
                }
            }
        }

        private static void PopulateAdjustmentRow(
            object[,] output,
            int firstRow,
            WorkbookEstimatePreview preview)
        {
            int row = preview.Plan.AdjustmentRow;
            if (row <= 0)
                return;
            WorkbookEstimateLine[] waterLines = preview.Plan.Lines
                .Where(line => string.Equals(line.GroupKey, "WATER", StringComparison.Ordinal))
                .OrderBy(line => line.TargetRow)
                .ToArray();
            if (waterLines.Length == 0)
                return;

            EstimateAppendixGroupResult current = preview.Result.Groups.Single(
                group => string.Equals(group.GroupKey, "WATER", StringComparison.Ordinal));
            EstimateAppendixGroupResult baseline = preview.BaselineResult.Groups.Single(
                group => string.Equals(group.GroupKey, "WATER", StringComparison.Ordinal));
            decimal[] currentTotals =
            {
                current.MaterialAmountVnd,
                current.LaborAmountVnd,
                current.MachineAmountVnd,
                current.AcceptedMaterialAmountVnd,
                current.AcceptedLaborAmountVnd,
                current.AcceptedMachineAmountVnd
            };
            decimal[] baselineTotals =
            {
                baseline.MaterialAmountVnd,
                baseline.LaborAmountVnd,
                baseline.MachineAmountVnd,
                baseline.AcceptedMaterialAmountVnd,
                baseline.AcceptedLaborAmountVnd,
                baseline.AcceptedMachineAmountVnd
            };
            int start = waterLines.First().TargetRow;
            int end = waterLines.Last().TargetRow;
            for (int index = 0; index < 6; index++)
            {
                int column = 9 + index;
                decimal delta = currentTotals[index] - baselineTotals[index];
                decimal multiplier;
                if (TryGetUniformMultiplier(preview, index % 3, out multiplier))
                {
                    decimal adjustment = multiplier - 1m;
                    if (adjustment == 0m)
                    {
                        Set(output, firstRow, row, column, 0d);
                    }
                    else
                    {
                        string name = ColumnName(column);
                        Set(output, firstRow, row, column,
                            "=SUM(" + name + start.ToString(CultureInfo.InvariantCulture) + ":" +
                            name + end.ToString(CultureInfo.InvariantCulture) + ")*" +
                            adjustment.ToString("0.############################", CultureInfo.InvariantCulture));
                    }
                }
                else
                {
                    Set(output, firstRow, row, column, Number(delta));
                }
            }
        }

        private static bool TryGetUniformMultiplier(
            WorkbookEstimatePreview preview,
            int component,
            out decimal multiplier)
        {
            multiplier = 1m;
            bool found = false;
            foreach (WorkbookEstimatePreviewLine line in preview.Lines.Where(
                item => string.Equals(item.Line.GroupKey, "WATER", StringComparison.Ordinal)))
            {
                decimal baseline = RateComponent(line.BaselineUnitRate, component);
                decimal current = RateComponent(line.UnitRate, component);
                if (baseline == 0m)
                {
                    if (current != 0m)
                        return false;
                    continue;
                }
                decimal candidate = current / baseline;
                if (!found)
                {
                    multiplier = candidate;
                    found = true;
                }
                else if (Math.Abs(candidate - multiplier) > 0.0000000001m)
                {
                    return false;
                }
            }
            return true;
        }

        private static decimal RateComponent(UnitRateCalculationResult rate, int component)
        {
            if (rate == null)
                return 0m;
            switch (component)
            {
                case 0: return rate.MaterialAmountVnd;
                case 1: return rate.LaborAmountVnd;
                case 2: return rate.MachineAmountVnd;
                default: throw new ArgumentOutOfRangeException(nameof(component));
            }
        }

        private static void PopulateGrandTotalRow(
            object[,] output,
            int firstRow,
            WorkbookEstimatePreview preview)
        {
            int row = preview.Plan.GrandTotalRow;
            if (row <= 0)
                throw new InvalidOperationException("Khong tim thay dong Tong cong cua phu luc.");
            int[] subtotalRows = preview.Plan.Lines
                .GroupBy(line => line.GroupKey, StringComparer.Ordinal)
                .Select(group => group.Min(line => line.TargetRow) - 1)
                .OrderBy(value => value)
                .ToArray();
            for (int column = 9; column <= 14; column++)
            {
                string name = ColumnName(column);
                string formula = "=" + string.Join("+", subtotalRows.Select(
                    subtotal => name + subtotal.ToString(CultureInfo.InvariantCulture)));
                Set(output, firstRow, row, column, formula);
            }
        }

        private static void VerifyGrandTotal(
            Excel.Worksheet worksheet,
            WorkbookEstimatePreview preview)
        {
            Excel.Range range = null;
            try
            {
                int row = preview.Plan.GrandTotalRow;
                range = worksheet.Range[
                    "I" + row.ToString(CultureInfo.InvariantCulture),
                    "N" + row.ToString(CultureInfo.InvariantCulture)];
                object values = range.Value2;
                decimal[] expected =
                {
                    preview.Result.MaterialAmountVnd,
                    preview.Result.LaborAmountVnd,
                    preview.Result.MachineAmountVnd,
                    preview.Result.AcceptedMaterialAmountVnd,
                    preview.Result.AcceptedLaborAmountVnd,
                    preview.Result.AcceptedMachineAmountVnd
                };
                for (int index = 0; index < expected.Length; index++)
                {
                    decimal actual = Convert.ToDecimal(
                        ((Array)values).GetValue(1, index + 1) ?? 0d,
                        CultureInfo.InvariantCulture);
                    if (Math.Abs(actual - expected[index]) > Tolerance)
                    {
                        throw new InvalidOperationException(
                            "Kiem tra tong phu luc that bai tai " + ColumnName(9 + index) + row +
                            ": expected=" + expected[index].ToString(CultureInfo.InvariantCulture) +
                            ", actual=" + actual.ToString(CultureInfo.InvariantCulture) + ".");
                    }
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

        private static void Set(
            object[,] values,
            int firstRow,
            int worksheetRow,
            int worksheetColumn,
            object value)
        {
            values[worksheetRow - firstRow, worksheetColumn - 6] = value;
        }

        private static double Number(decimal value)
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        private static string ColumnName(int column)
        {
            const string names = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            if (column < 1 || column > names.Length)
                throw new ArgumentOutOfRangeException(nameof(column));
            return names[column - 1].ToString();
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
