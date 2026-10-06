using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelAddIn1.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public static partial class RandomDaodat
    {
        private const string SpecialSourceSerialNumber = "KEY_STT";

        private static Dictionary<string, string> BuildOutputColumnMap(DaodatRunOptions options)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DaodatColumnMapping mapping in options.ColumnMappings ?? new List<DaodatColumnMapping>())
            {
                if (string.IsNullOrWhiteSpace(mapping.Key) || string.IsNullOrWhiteSpace(mapping.OutputColumn))
                    continue;

                map[NormalizeKey(mapping.Key)] = ExcelColumnAddress.Normalize(mapping.OutputColumn);
            }

            return map;
        }

        private static Dictionary<string, string> BuildOutputFormatMap(DaodatRunOptions options)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DaodatColumnMapping mapping in options.ColumnMappings ?? new List<DaodatColumnMapping>())
            {
                if (string.IsNullOrWhiteSpace(mapping.Key) || string.IsNullOrWhiteSpace(mapping.OutputColumn))
                    continue;

                map[NormalizeKey(mapping.Key)] = GetMappingFormatLocal(mapping);
            }

            return map;
        }

        private static string NormalizeKey(string value)
        {
            return (value ?? string.Empty).Trim().ToUpperInvariant();
        }

        private static bool IsGeneratedOutputKey(string value)
        {
            string key = NormalizeKey(value);
            switch (key)
            {
                case "TH3":
                case "V3":
                case "D1_3":
                case "R1_3":
                case "D2_3":
                case "R2_3":
                case "H_3":
                case "TH5":
                case "V5":
                case "D1_5":
                case "R1_5":
                case "D2_5":
                case "R2_5":
                case "H_5":
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsSerialNumberMapping(DaodatColumnMapping mapping)
        {
            return mapping != null && IsSerialNumberSourceKey(mapping.SourceColumn);
        }

        private static bool IsSerialNumberSourceKey(string value)
        {
            return string.Equals(NormalizeKey(value), SpecialSourceSerialNumber, StringComparison.OrdinalIgnoreCase);
        }

        private static string EscapeSheetName(string sheetName)
        {
            return (sheetName ?? string.Empty).Replace("'", "''");
        }

        private static Excel.Range GetColumnRange(Excel.Worksheet worksheet, int startRow, int column, int rowCount)
        {
            return worksheet.Range[
                worksheet.Cells[startRow, column],
                worksheet.Cells[startRow + rowCount - 1, column]];
        }

        private static void SetColumnValues(Excel.Worksheet worksheet, int startRow, int column, object[] values, string numberFormatLocal = null)
        {
            if (values == null || values.Length == 0)
                return;

            object[,] data = new object[values.Length, 1];
            for (int i = 0; i < values.Length; i++)
                data[i, 0] = values[i];

            Excel.Range range = null;
            try
            {
                range = GetColumnRange(worksheet, startRow, column, values.Length);
                range.Value2 = data;
                ApplyNumberFormatLocal(range, numberFormatLocal);
            }
            finally
            {
                ReleaseComObject(range);
            }
        }

        private static void ApplyNumberFormatLocal(Excel.Range range, string numberFormatLocal)
        {
            if (range == null)
                return;

            range.NumberFormatLocal = string.IsNullOrWhiteSpace(numberFormatLocal)
                ? "General"
                : numberFormatLocal;
        }

        private static string GetOutputFormatLocal(Dictionary<string, string> outputFormats, string key)
        {
            if (outputFormats != null
                && outputFormats.TryGetValue(key, out string formatLocal)
                && !string.IsNullOrWhiteSpace(formatLocal))
            {
                return formatLocal;
            }

            return GetDefaultFormatLocal(key, string.Empty);
        }

        private static string GetMappingFormatLocal(DaodatColumnMapping mapping)
        {
            if (mapping == null)
                return "General";

            return string.IsNullOrWhiteSpace(mapping.FormatLocal)
                ? GetDefaultFormatLocal(mapping.Key, mapping.SourceColumn)
                : mapping.FormatLocal.Trim();
        }

        private static string GetDefaultFormatLocal(string key, string sourceColumn)
        {
            if (IsSerialNumberSourceKey(sourceColumn))
                return "0";

            switch (NormalizeKey(key))
            {
                case "TH3":
                case "TH5":
                    return "0";
                case "V3":
                case "V5":
                case "D1_3":
                case "R1_3":
                case "D2_3":
                case "R2_3":
                case "D1_5":
                case "R1_5":
                case "D2_5":
                case "R2_5":
                    return DecimalFormatLocal(2);
                case "H_3":
                case "H_5":
                    return DecimalFormatLocal(4);
                default:
                    return "General";
            }
        }

        private static string DecimalFormatLocal(int decimalPlaces)
        {
            if (decimalPlaces <= 0)
                return "0";

            return "0"
                + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
                + new string('0', decimalPlaces);
        }

        private static void InsertWholeRows(Excel.Worksheet worksheet, int startRow, int rowCount)
        {
            if (worksheet == null || rowCount <= 0)
                return;

            Excel.Range rows = worksheet.Range[
                worksheet.Cells[startRow, 1],
                worksheet.Cells[startRow + rowCount - 1, 1]].EntireRow;
            rows.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
        }

        private static string BuildVolumeFormulaR1C1(int vCol, int d1Col, int r1Col, int d2Col, int r2Col, int hCol)
        {
            string d1 = ToR1C1Offset(d1Col - vCol);
            string r1 = ToR1C1Offset(r1Col - vCol);
            string d2 = ToR1C1Offset(d2Col - vCol);
            string r2 = ToR1C1Offset(r2Col - vCol);
            string h = ToR1C1Offset(hCol - vCol);
            return $"=ROUND((({d1}*{r1})+({d2}*{r2})+SQRT(({d1}*{r1})*({d2}*{r2})))*{h}/3,2)";
        }

        private static string ToR1C1Offset(int offset)
        {
            if (offset == 0)
                return "RC";

            return offset > 0 ? $"RC[{offset}]" : $"RC[{offset}]";
        }

        private static int ReadIntFromMappedColumn(Excel.Worksheet sheet, int row, Dictionary<string, string> columns, string key)
        {
            double value = ReadDoubleFromMappedColumn(sheet, row, columns, key);
            return value > 0 ? Convert.ToInt32(Math.Round(value, 0)) : 0;
        }

        private static double ReadDoubleFromMappedColumn(Excel.Worksheet sheet, int row, Dictionary<string, string> columns, string key)
        {
            if (columns == null || !columns.TryGetValue(key, out string colLetter) || string.IsNullOrWhiteSpace(colLetter))
                return 0;

            Excel.Range cell = sheet.Cells[row, ColumnLetterToNumber(colLetter)] as Excel.Range;
            object value = cell?.Value2;
            if (value == null)
                return 0;

            if (value is double doubleValue)
                return doubleValue;
            if (value is int intValue)
                return intValue;
            if (value is decimal decimalValue)
                return Convert.ToDouble(decimalValue);

            string text = value.ToString();
            if (double.TryParse(text, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.CurrentCulture, out double currentNumber))
                return currentNumber;
            if (double.TryParse(text, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, System.Globalization.CultureInfo.InvariantCulture, out double invariantNumber))
                return invariantNumber;

            return 0;
        }

        private static Excel.Range GetRangeFromWorkbookAddress(Excel.Workbook workbook, string address)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(address))
                return null;

            string text = address.Trim();
            int bangIndex = text.LastIndexOf('!');
            try
            {
                if (bangIndex > 0)
                {
                    string sheetName = text.Substring(0, bangIndex).Trim('\'');
                    string rangeAddress = text.Substring(bangIndex + 1);
                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                            return sheet.Range[rangeAddress] as Excel.Range;
                    }

                    return null;
                }

                return Globals.ThisAddIn.Application.ActiveSheet.Range[text] as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        private static int ColumnLetterToNumber(string column)
        {
            return ExcelColumnAddress.ToNumber(column);
        }

        private static string BuildFailureMessage(Excel.Range cell, double targetSum, int count, DaodatRandomSettings settings)
        {
            return DaodatFeasibility.BuildFailureMessage(
                "o " + cell.Address[false, false],
                targetSum,
                count,
                settings.VolumeMin,
                settings.VolumeMax,
                settings.MaxAttempts);
        }

        private static object[,] CloneRangeMatrix(object source, int rows, int columns)
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

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private static string BuildFailureMessage(string label, double targetSum, int count, DaodatRandomSettings settings)
        {
            return DaodatFeasibility.BuildFailureMessage(
                label,
                targetSum,
                count,
                settings.VolumeMin,
                settings.VolumeMax,
                settings.MaxAttempts);
        }

        private static void EnsureTargetFeasible(string label, int count, double targetSum, DaodatRandomSettings settings)
        {
            DaodatFeasibilityResult feasibility = DaodatFeasibility.Analyze(count, targetSum, settings.VolumeMin, settings.VolumeMax);
            if (!feasibility.CanGenerate)
                throw new InvalidOperationException(BuildFailureMessage(label, targetSum, count, settings));
        }

        private static string GetUniqueSheetName(Excel.Workbook workbook, string prefix)
        {
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string baseName = $"{prefix}_{stamp}";
            string name = baseName;
            int index = 1;

            while (SheetExists(workbook, name))
            {
                name = $"{baseName}_{index}";
                index++;
            }

            return name;
        }

        private static bool SheetExists(Excel.Workbook workbook, string name)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private static void LinkBackToSource(
            Excel.Range sourceRange,
            Excel.Worksheet resultSheet,
            int nCol,
            List<(int sourceRow, Excel.Range resultCell, Excel.Range resultCountCell)> linkMappings)
        {
            foreach (var mapping in linkMappings)
            {
                Excel.Range sourceCell = sourceRange.Worksheet.Cells[mapping.sourceRow, nCol] as Excel.Range;
                if (sourceCell?.Value2 == null)
                    continue;

                Excel.Range sourceCountCell = sourceRange.Worksheet.Cells[mapping.sourceRow, nCol - 1] as Excel.Range;

                sourceCell.Formula = $"='{EscapeSheetName(resultSheet.Name)}'!{mapping.resultCell.Address[false, false]}";
                sourceCountCell.Formula = $"='{EscapeSheetName(resultSheet.Name)}'!{mapping.resultCountCell.Address[false, false]}";
            }
        }
    }
}
