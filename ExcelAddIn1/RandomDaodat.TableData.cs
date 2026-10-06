using ExcelAddIn1.Funtion;
using ExcelAddIn1.Core;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public static partial class RandomDaodat
    {
        private static void DaodatFromTableData(DaodatRunOptions runOptions)
        {
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Excel.Worksheet createdWorksheet = null;
            bool completed = false;
            ExcelWriteContext writeContext = null;

            try
            {
                SplashManager.ShowForm("Vui long cho!", "Dang xu ly du lieu...");
                writeContext = new ExcelWriteContext(xlApp);

                Excel.Range dataRange = GetRangeFromWorkbookAddress(xlApp.ActiveWorkbook, runOptions.TableDataAddress);
                if (dataRange == null)
                    throw new InvalidOperationException("Range Table Data khong hop le.");

                if (runOptions.OutputStartRow <= 0)
                    throw new InvalidOperationException("Hang trong dao dap phai lon hon 0.");

                Excel.Worksheet dataSheet = dataRange.Worksheet as Excel.Worksheet;
                Excel.Worksheet outputSheet = ResolveOutputSheet(xlApp.ActiveWorkbook, dataSheet, runOptions, out createdWorksheet);
                Dictionary<string, string> outputColumns = BuildOutputColumnMap(runOptions);
                Dictionary<string, string> outputFormats = BuildOutputFormatMap(runOptions);

                DaodatRandomSettings settings3m = GetSettings(false);
                DaodatRandomSettings settings5m = GetSettings(true);
                var rowPlans = BuildTableRowPlans(dataRange, runOptions, settings3m, settings5m);
                int totalRows = rowPlans.Sum(x => x.BlockRows);
                int headerRowCount = 1;

                if (totalRows <= 0)
                    throw new InvalidOperationException("Khong co dong du lieu hop le de tinh.");

                if (!runOptions.CreateNewSheet)
                {
                    InsertWholeRows(outputSheet, runOptions.OutputStartRow, totalRows + headerRowCount);
                }

                UnmergeOutputWriteArea(outputSheet, runOptions.OutputStartRow, totalRows + headerRowCount, runOptions.ColumnMappings);
                WriteTableHeaderRow(outputSheet, runOptions.OutputStartRow, runOptions.ColumnMappings);

                int currentRow = runOptions.OutputStartRow + headerRowCount;
                int summaryIndex = 1;
                foreach (var plan in rowPlans)
                {
                    SplashManager.SetDescription("Dang ghi block " + summaryIndex.ToString() + "/" + rowPlans.Count.ToString() + "...");
                    WriteSpecialSourceColumns(outputSheet, currentRow, summaryIndex, runOptions.ColumnMappings);

                    if (plan.Results3m != null)
                        WriteGeneratedResultsToColumns(outputSheet, currentRow, plan.BlockRows, plan.Results3m, outputColumns, outputFormats, false, xlApp);

                    if (runOptions.ProjectIncludes5m && plan.Results5m != null)
                        WriteGeneratedResultsToColumns(outputSheet, currentRow, plan.BlockRows, plan.Results5m, outputColumns, outputFormats, true, xlApp);

                    LinkAuxiliarySourceColumns(dataSheet, outputSheet, plan.SourceRow, currentRow, runOptions.ColumnMappings);

                    if (runOptions.LinkBack)
                        LinkSummaryBackToSource(dataSheet, outputSheet, plan.SourceRow, currentRow, plan, runOptions, outputColumns);

                    currentRow += plan.BlockRows;
                    summaryIndex++;
                }

                outputSheet.Activate();
                completed = true;
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "RandomDaodat.DaodatFromTableData", new Dictionary<string, string>
                {
                    ["TableDataAddress"] = runOptions.TableDataAddress ?? string.Empty,
                    ["OutputStartRow"] = runOptions.OutputStartRow.ToString(),
                    ["CreateNewSheet"] = runOptions.CreateNewSheet.ToString(),
                    ["ExistingSheetName"] = runOptions.ExistingSheetName ?? string.Empty
                });

                if (!completed && createdWorksheet != null)
                    createdWorksheet.Delete();

                MessageBox.Show(ex.Message + Environment.NewLine + "Log: " + RuntimeLogger.LogFilePath, "Khong tao duoc ho dao", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                writeContext?.Dispose();
                SplashManager.CloseForm();
            }
        }

        private static void UnmergeOutputWriteArea(
            Excel.Worksheet outputSheet,
            int startRow,
            int rowCount,
            List<DaodatColumnMapping> mappings)
        {
            List<int> outputColumns = new List<int>();
            foreach (DaodatColumnMapping mapping in mappings ?? new List<DaodatColumnMapping>())
            {
                if (string.IsNullOrWhiteSpace(mapping.OutputColumn))
                    continue;

                outputColumns.Add(ColumnLetterToNumber(mapping.OutputColumn));
            }

            if (outputColumns.Count == 0)
                return;

            int firstColumn = outputColumns.Min();
            int lastColumn = outputColumns.Max();
            Excel.Range writeArea = null;
            try
            {
                writeArea = outputSheet.Range[
                    outputSheet.Cells[startRow, firstColumn],
                    outputSheet.Cells[startRow + rowCount - 1, lastColumn]];
                if (writeArea.MergeCells)
                    writeArea.UnMerge();
            }
            finally
            {
                ReleaseComObject(writeArea);
            }
        }

        private static void WriteTableHeaderRow(
            Excel.Worksheet outputSheet,
            int headerRow,
            List<DaodatColumnMapping> mappings)
        {
            List<DaodatColumnMapping> active = (mappings ?? new List<DaodatColumnMapping>())
                .Where(mapping => !string.IsNullOrWhiteSpace(mapping.Key) &&
                    !string.IsNullOrWhiteSpace(mapping.OutputColumn))
                .ToList();
            if (active.Count == 0)
                return;
            int firstColumn = active.Min(mapping => ColumnLetterToNumber(mapping.OutputColumn));
            int lastColumn = active.Max(mapping => ColumnLetterToNumber(mapping.OutputColumn));
            Excel.Range rowRange = null;
            try
            {
                rowRange = outputSheet.Range[
                    outputSheet.Cells[headerRow, firstColumn],
                    outputSheet.Cells[headerRow, lastColumn]];
                object[,] values = CloneRangeMatrix(rowRange.Formula, 1, lastColumn - firstColumn + 1);
                foreach (DaodatColumnMapping mapping in active)
                    values[0, ColumnLetterToNumber(mapping.OutputColumn) - firstColumn] = mapping.Key.Trim();
                rowRange.Formula = values;
                foreach (DaodatColumnMapping mapping in active)
                {
                    Excel.Range cell = null;
                    try
                    {
                        cell = outputSheet.Cells[headerRow, ColumnLetterToNumber(mapping.OutputColumn)] as Excel.Range;
                        cell.Font.Bold = true;
                    }
                    finally
                    {
                        ReleaseComObject(cell);
                    }
                }
            }
            finally
            {
                ReleaseComObject(rowRange);
            }
        }

        private static void WriteSpecialSourceColumns(
            Excel.Worksheet outputSheet,
            int summaryRow,
            int summaryIndex,
            List<DaodatColumnMapping> mappings)
        {
            foreach (DaodatColumnMapping mapping in mappings ?? new List<DaodatColumnMapping>())
            {
                if (!IsSerialNumberMapping(mapping) || string.IsNullOrWhiteSpace(mapping.OutputColumn))
                    continue;

                int outputCol = ColumnLetterToNumber(mapping.OutputColumn);
                Excel.Range cell = null;
                try
                {
                    cell = outputSheet.Cells[summaryRow, outputCol] as Excel.Range;
                    cell.Value2 = summaryIndex;
                    ApplyNumberFormatLocal(cell, GetMappingFormatLocal(mapping));
                    cell.Font.Bold = true;
                }
                finally
                {
                    ReleaseComObject(cell);
                }
            }
        }

        private class TableRowPlan
        {
            public int SourceRow { get; set; }
            public int BlockRows { get; set; }
            public List<(double d1, double r1, double d2, double r2, double H)> Results3m { get; set; }
            public List<(double d1, double r1, double d2, double r2, double H)> Results5m { get; set; }
        }

        private static List<TableRowPlan> BuildTableRowPlans(
            Excel.Range dataRange,
            DaodatRunOptions runOptions,
            DaodatRandomSettings settings3m,
            DaodatRandomSettings settings5m)
        {
            var plans = new List<TableRowPlan>();
            int firstRow = dataRange.Row;
            int rowCount = dataRange.Rows.Count;

            for (int i = 0; i < rowCount; i++)
            {
                int row = firstRow + i;
                SplashManager.SetDescription("Dang tinh dong data " + (i + 1).ToString() + "/" + rowCount.ToString() + "...");
                int th3 = ReadIntFromMappedColumn(dataRange.Worksheet, row, runOptions.SourceDataColumns, "TH3");
                double v3 = ReadDoubleFromMappedColumn(dataRange.Worksheet, row, runOptions.SourceDataColumns, "V3");
                int th5 = runOptions.ProjectIncludes5m ? ReadIntFromMappedColumn(dataRange.Worksheet, row, runOptions.SourceDataColumns, "TH5") : 0;
                double v5 = runOptions.ProjectIncludes5m ? ReadDoubleFromMappedColumn(dataRange.Worksheet, row, runOptions.SourceDataColumns, "V5") : 0;

                if (th3 <= 0 && th5 <= 0)
                    continue;

                if (th3 > 0 && v3 <= 0)
                    throw new InvalidOperationException($"Dong {row}: TH3 > 0 nhung V3 khong hop le.");

                if (runOptions.ProjectIncludes5m && th5 > 0 && v5 <= 0)
                    throw new InvalidOperationException($"Dong {row}: TH5 > 0 nhung V5 khong hop le.");

                var plan = new TableRowPlan
                {
                    SourceRow = row,
                    BlockRows = runOptions.ProjectIncludes5m ? Math.Max(th3, th5) + 1 : th3 + 1
                };

                if (th3 > 0 && v3 > 0)
                    plan.Results3m = GenerateResultsForTarget(th3, v3, settings3m, $"dong {row} V3");

                if (runOptions.ProjectIncludes5m && th5 > 0 && v5 > 0)
                    plan.Results5m = GenerateResultsForTarget(th5, v5, settings5m, $"dong {row} V5");

                plans.Add(plan);
            }

            return plans;
        }

        private static void WriteGeneratedResultsToColumns(
            Excel.Worksheet worksheet,
            int startRow,
            int blockRows,
            List<(double d1, double r1, double d2, double r2, double H)> results,
            Dictionary<string, string> outputColumns,
            Dictionary<string, string> outputFormats,
            bool is5m,
            Excel.Application xlApp)
        {
            string suffix = is5m ? "_5" : "_3";
            string thKey = is5m ? "TH5" : "TH3";
            string vKey = is5m ? "V5" : "V3";
            bool hasThCol = TryGetOutputColumn(outputColumns, thKey, out int thCol);
            bool hasVCol = TryGetOutputColumn(outputColumns, vKey, out int vCol);
            bool hasD1Col = TryGetOutputColumn(outputColumns, "D1" + suffix, out int d1Col);
            bool hasR1Col = TryGetOutputColumn(outputColumns, "R1" + suffix, out int r1Col);
            bool hasD2Col = TryGetOutputColumn(outputColumns, "D2" + suffix, out int d2Col);
            bool hasR2Col = TryGetOutputColumn(outputColumns, "R2" + suffix, out int r2Col);
            bool hasHCol = TryGetOutputColumn(outputColumns, "H" + suffix, out int hCol);
            int detailStartRow = startRow + 1;
            string thFormat = GetOutputFormatLocal(outputFormats, thKey);
            string vFormat = GetOutputFormatLocal(outputFormats, vKey);
            string d1Format = GetOutputFormatLocal(outputFormats, "D1" + suffix);
            string r1Format = GetOutputFormatLocal(outputFormats, "R1" + suffix);
            string d2Format = GetOutputFormatLocal(outputFormats, "D2" + suffix);
            string r2Format = GetOutputFormatLocal(outputFormats, "R2" + suffix);
            string hFormat = GetOutputFormatLocal(outputFormats, "H" + suffix);

            if (hasThCol)
                SetColumnValues(worksheet, detailStartRow, thCol, results.Select(_ => (object)1).ToArray(), thFormat);
            if (hasD1Col)
                SetColumnValues(worksheet, detailStartRow, d1Col, results.Select(x => (object)Math.Round(x.d1, 2)).ToArray(), d1Format);
            if (hasR1Col)
                SetColumnValues(worksheet, detailStartRow, r1Col, results.Select(x => (object)Math.Round(x.r1, 2)).ToArray(), r1Format);
            if (hasD2Col)
                SetColumnValues(worksheet, detailStartRow, d2Col, results.Select(x => (object)Math.Round(x.d2, 2)).ToArray(), d2Format);
            if (hasR2Col)
                SetColumnValues(worksheet, detailStartRow, r2Col, results.Select(x => (object)Math.Round(x.r2, 2)).ToArray(), r2Format);
            if (hasHCol)
                SetColumnValues(worksheet, detailStartRow, hCol, results.Select(x => (object)Math.Round(x.H, 4)).ToArray(), hFormat);

            if (hasVCol)
            {
                Excel.Range volumeRange = null;
                try
                {
                    volumeRange = GetColumnRange(worksheet, detailStartRow, vCol, results.Count);
                    if (hasD1Col && hasR1Col && hasD2Col && hasR2Col && hasHCol)
                    {
                        volumeRange.FormulaR1C1 = BuildVolumeFormulaR1C1(vCol, d1Col, r1Col, d2Col, r2Col, hCol);
                        ApplyNumberFormatLocal(volumeRange, vFormat);
                    }
                    else
                    {
                        SetColumnValues(
                            worksheet,
                            detailStartRow,
                            vCol,
                            results.Select(x => (object)Math.Round(DaodatRandomSettings.CalculateVolume(x.d1, x.r1, x.d2, x.r2, x.H), 2)).ToArray(),
                            vFormat);
                    }
                }
                finally
                {
                    ReleaseComObject(volumeRange);
                }
            }

            if (hasThCol)
                WriteSummaryFormula(worksheet, startRow, blockRows, thCol, thFormat);
            if (hasVCol)
                WriteSummaryFormula(worksheet, startRow, blockRows, vCol, vFormat);
        }

        private static bool TryGetOutputColumn(Dictionary<string, string> outputColumns, string key, out int column)
        {
            column = 0;
            if (outputColumns == null || !outputColumns.TryGetValue(key, out string columnText) || string.IsNullOrWhiteSpace(columnText))
                return false;

            column = ColumnLetterToNumber(columnText);
            return true;
        }

        private static void WriteSummaryFormula(Excel.Worksheet worksheet, int startRow, int blockRows, int column, string numberFormatLocal)
        {
            Excel.Range summaryCell = null;
            try
            {
                summaryCell = worksheet.Cells[startRow, column] as Excel.Range;
                summaryCell.FormulaR1C1 = $"=SUM(R[1]C:R[{blockRows - 1}]C)";
                ApplyNumberFormatLocal(summaryCell, numberFormatLocal);
                summaryCell.Font.Bold = true;
            }
            finally
            {
                ReleaseComObject(summaryCell);
            }
        }

        private static void LinkAuxiliarySourceColumns(
            Excel.Worksheet dataSheet,
            Excel.Worksheet outputSheet,
            int sourceRow,
            int outputRow,
            List<DaodatColumnMapping> mappings)
        {
            List<DaodatColumnMapping> active = (mappings ?? new List<DaodatColumnMapping>())
                .Where(mapping => !string.IsNullOrWhiteSpace(mapping.SourceColumn) &&
                    !string.IsNullOrWhiteSpace(mapping.OutputColumn) &&
                    !IsGeneratedOutputKey(mapping.Key) &&
                    !IsSerialNumberMapping(mapping))
                .ToList();
            if (active.Count == 0)
                return;
            int firstColumn = active.Min(mapping => ColumnLetterToNumber(mapping.OutputColumn));
            int lastColumn = active.Max(mapping => ColumnLetterToNumber(mapping.OutputColumn));
            Excel.Range outputRange = null;
            try
            {
                outputRange = outputSheet.Range[
                    outputSheet.Cells[outputRow, firstColumn],
                    outputSheet.Cells[outputRow, lastColumn]];
                object[,] formulas = CloneRangeMatrix(outputRange.Formula, 1, lastColumn - firstColumn + 1);
                foreach (DaodatColumnMapping mapping in active)
                {
                    int sourceCol = ColumnLetterToNumber(mapping.SourceColumn);
                    int outputCol = ColumnLetterToNumber(mapping.OutputColumn);
                    formulas[0, outputCol - firstColumn] = "='" + EscapeSheetName(dataSheet.Name) +
                        "'!" + ExcelColumnAddress.Normalize(mapping.SourceColumn) + sourceRow;
                }
                outputRange.Formula = formulas;
                foreach (DaodatColumnMapping mapping in active)
                {
                    Excel.Range outputCell = null;
                    try
                    {
                        outputCell = outputSheet.Cells[outputRow, ColumnLetterToNumber(mapping.OutputColumn)] as Excel.Range;
                        ApplyNumberFormatLocal(outputCell, GetMappingFormatLocal(mapping));
                    }
                    finally
                    {
                        ReleaseComObject(outputCell);
                    }
                }
            }
            finally
            {
                ReleaseComObject(outputRange);
            }
        }

        private static void LinkSummaryBackToSource(
            Excel.Worksheet dataSheet,
            Excel.Worksheet outputSheet,
            int sourceRow,
            int outputRow,
            TableRowPlan plan,
            DaodatRunOptions options,
            Dictionary<string, string> outputColumns)
        {
            if (options.SourceDataColumns == null || outputColumns == null)
                return;
            var keys = new List<string>();
            if (plan.Results3m != null)
                keys.AddRange(new[] { "TH3", "V3" });
            if (options.ProjectIncludes5m && plan.Results5m != null)
                keys.AddRange(new[] { "TH5", "V5" });
            var pairs = keys
                .Where(key => options.SourceDataColumns.TryGetValue(key, out string source) &&
                    !string.IsNullOrWhiteSpace(source) &&
                    outputColumns.TryGetValue(key, out string output) &&
                    !string.IsNullOrWhiteSpace(output))
                .Select(key => new
                {
                    SourceColumn = ColumnLetterToNumber(options.SourceDataColumns[key]),
                    OutputColumn = ColumnLetterToNumber(outputColumns[key])
                })
                .ToList();
            if (pairs.Count == 0)
                return;
            int firstColumn = pairs.Min(pair => pair.SourceColumn);
            int lastColumn = pairs.Max(pair => pair.SourceColumn);
            Excel.Range sourceRange = null;
            try
            {
                sourceRange = dataSheet.Range[
                    dataSheet.Cells[sourceRow, firstColumn],
                    dataSheet.Cells[sourceRow, lastColumn]];
                object[,] formulas = CloneRangeMatrix(sourceRange.Formula, 1, lastColumn - firstColumn + 1);
                foreach (var pair in pairs)
                {
                    formulas[0, pair.SourceColumn - firstColumn] = "='" + EscapeSheetName(outputSheet.Name) +
                        "'!" + ExcelColumnAddress.ToLetters(pair.OutputColumn) + outputRow;
                }
                sourceRange.Formula = formulas;
            }
            finally
            {
                ReleaseComObject(sourceRange);
            }
        }

        private static Excel.Worksheet ResolveOutputSheet(
            Excel.Workbook workbook,
            Excel.Worksheet dataSheet,
            DaodatRunOptions options,
            out Excel.Worksheet createdWorksheet)
        {
            createdWorksheet = null;
            if (options.CreateNewSheet)
            {
                Excel.Worksheet sheet = workbook.Worksheets.Add(After: dataSheet) as Excel.Worksheet;
                sheet.Name = GetUniqueSheetName(workbook, string.IsNullOrWhiteSpace(options.NewSheetName) ? "HoDao" : options.NewSheetName.Trim());
                createdWorksheet = sheet;
                return sheet;
            }

            if (string.IsNullOrWhiteSpace(options.ExistingSheetName))
                throw new InvalidOperationException("Vui long chon sheet hien co.");

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, options.ExistingSheetName, StringComparison.OrdinalIgnoreCase))
                    return sheet;
            }

            throw new InvalidOperationException("Khong tim thay sheet: " + options.ExistingSheetName);
        }
    }
}
