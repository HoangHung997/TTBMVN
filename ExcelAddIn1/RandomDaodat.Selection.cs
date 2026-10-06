using ExcelAddIn1.Funtion;
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
        public static void Daodat(bool is5m, bool isRowNew)
        {
            DaodatRunOptions savedOptions = GetRunOptions(is5m);
            savedOptions.Is5m = is5m;
            savedOptions.InsertRows = isRowNew;
            Daodat(savedOptions);
        }

        public static void Daodat(DaodatRunOptions runOptions)

        {
            if (runOptions == null)
                throw new ArgumentNullException(nameof(runOptions));

            if (runOptions.HasTableData)
            {
                DaodatFromTableData(runOptions);
                return;
            }

            //SplashScreenManager.ShowForm(typeof(FrmWaiting));
            var xlApp = Globals.ThisAddIn.Application;
            var sourceWorksheet = xlApp.ActiveSheet as Excel.Worksheet;
            var selectedRange = xlApp.Selection as Excel.Range;

            if (selectedRange == null)
            {
                MessageBox.Show("Vui long chon vung du lieu truoc khi chay.", "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DaodatRandomSettings settings = GetSettings(runOptions.Is5m);
            settings.Validate();

            Excel.Worksheet createdWorksheet = null;
            bool completed = false;
            ExcelWriteContext writeContext = null;
            try
            {
                SplashManager.ShowForm("Vui long cho!", "Dang xu ly du lieu...");
                writeContext = new ExcelWriteContext(xlApp);
                if (!IsSingleColumnSelection(selectedRange))
                {
                    MessageBox.Show("Chi duoc chon cac o trong 1 cot, vui long chon lai", "Canh bao", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Excel.Worksheet worksheet = sourceWorksheet;
                Excel.Range processingRange = selectedRange;
                Excel.Range outputStartCell = null;
                Excel.Worksheet linkedSourceWorksheet = null;
                Excel.Range linkedSourceRange = null;

                if (runOptions.CreateNewSheet)
                {
                    sourceWorksheet.Copy(After: sourceWorksheet);
                    worksheet = xlApp.ActiveSheet as Excel.Worksheet;
                    worksheet.Name = GetUniqueSheetName(xlApp.ActiveWorkbook, "HoDao");
                    createdWorksheet = worksheet;
                    processingRange = worksheet.Range[selectedRange.Address];
                    linkedSourceWorksheet = sourceWorksheet;
                    linkedSourceRange = selectedRange;
                }

                if (runOptions.UseOutputStartCell)
                {
                    outputStartCell = GetRangeFromAddress(worksheet, runOptions.OutputStartAddress);
                    if (outputStartCell == null)
                        throw new InvalidOperationException("Cell bat dau xuat khong hop le.");
                }

                int nCol = selectedRange.Column;
                int firstRow = processingRange.Row;
                int lastRow = firstRow + processingRange.Rows.Count - 1;

                // Lưu các ô và giá trị để xử lý tránh thao tác Excel nhiều lần
                var cellsToProcess = new List<(Excel.Range cell, int N, double targetSum)>();
                for (int i = lastRow; i >= firstRow; i--)
                {
                    var cell = worksheet.Cells[i, nCol] as Excel.Range;
                    if (cell.Value2 == null) continue;

                    if (!double.TryParse(cell.Value2.ToString(), out double targetSum) || targetSum == 0)
                        continue;

                    var cellLeft = cell.Offset[0, -1] as Excel.Range;
                    if (cellLeft == null || cellLeft.Value2 == null) continue;

                    if (!int.TryParse(cellLeft.Value2.ToString(), out int N) || N <= 0)
                        continue;

                    cellsToProcess.Add((cell, N, targetSum));
                }

                // Định dạng số
                string numberFormat2 = DecimalFormatLocal(2);
                string numberFormatH = DecimalFormatLocal(4);
                string numberFormat4 = "0";
                var previewResults = new List<(Excel.Range cell, int sourceRow, List<(double d1, double r1, double d2, double r2, double H)> results)>();

                foreach (var (cell, N, targetSum) in cellsToProcess)
                {
                    SplashManager.SetDescription("Dang tinh " + cell.Address[false, false] + "...");
                    List<(double d1, double r1, double d2, double r2, double H)> results =
                        GenerateResultsForTarget(N, targetSum, settings, "o " + cell.Address[false, false]);

                    previewResults.Add((cell, cell.Row, results));
                }

                var linkMappings = new List<(int sourceRow, Excel.Range resultCell, Excel.Range resultCountCell)>();
                int outputOffset = 0;
                foreach (var item in previewResults)
                {
                    SplashManager.SetDescription("Dang ghi ket qua dong " + item.sourceRow.ToString() + "...");
                    Excel.Range targetSummaryCell = outputStartCell == null
                        ? item.cell
                        : outputStartCell.Offset[outputOffset, 0] as Excel.Range;

                    WriteResultsToExcel(targetSummaryCell, item.results, numberFormat2, numberFormatH, numberFormat4, xlApp, runOptions.InsertRows);
                    linkMappings.Add((item.sourceRow, targetSummaryCell, targetSummaryCell.Offset[0, -1]));

                    if (outputStartCell != null)
                        outputOffset += item.results.Count + 1;
                }

                if (runOptions.CreateNewSheet && runOptions.LinkBack && linkedSourceWorksheet != null && linkedSourceRange != null)
                {
                    LinkBackToSource(linkedSourceRange, worksheet, nCol, linkMappings);
                }

                if (runOptions.CreateNewSheet)
                    worksheet.Activate();

                completed = true;
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "RandomDaodat.Daodat selection", new Dictionary<string, string>
                {
                    ["Is5m"] = runOptions.Is5m.ToString(),
                    ["CreateNewSheet"] = runOptions.CreateNewSheet.ToString(),
                    ["UseOutputStartCell"] = runOptions.UseOutputStartCell.ToString()
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

        private static void WriteResultsToExcel(
            Excel.Range cell,
            List<(double d1, double r1, double d2, double r2, double H)> results,
            string numberFormat2,
            string numberFormatH,
            string numberFormat4,
            Microsoft.Office.Interop.Excel.Application xlApp,
            bool isRowNew)
        {
            if (isRowNew)
            {
                InsertWholeRows(cell.Worksheet, cell.Row + 1, results.Count);
            }

            int startRow = cell.Row + 1;
            int countCol = cell.Column - 1;
            int vCol = cell.Column;
            int d1Col = cell.Column + 1;
            int r1Col = cell.Column + 2;
            int d2Col = cell.Column + 3;
            int r2Col = cell.Column + 4;
            int hCol = cell.Column + 5;

            Excel.Range writeRange = null;
            Excel.Range summaryRange = null;
            Excel.Range countRange = null;
            Excel.Range volumeRange = null;
            Excel.Range d1Range = null;
            Excel.Range r1Range = null;
            Excel.Range d2Range = null;
            Excel.Range r2Range = null;
            Excel.Range hRange = null;
            try
            {
                writeRange = cell.Worksheet.Range[
                    cell.Worksheet.Cells[cell.Row, countCol],
                    cell.Worksheet.Cells[cell.Row + results.Count, hCol]];
                object[,] output = CloneRangeMatrix(writeRange.FormulaR1C1, results.Count + 1, 7);
                output[0, 0] = $"=SUM(R[1]C:R[{results.Count}]C)";
                output[0, 1] = $"=SUM(R[1]C:R[{results.Count}]C)";
                string volumeFormula = BuildVolumeFormulaR1C1(vCol, d1Col, r1Col, d2Col, r2Col, hCol);
                for (int index = 0; index < results.Count; index++)
                {
                    var item = results[index];
                    output[index + 1, 0] = 1;
                    output[index + 1, 1] = volumeFormula;
                    output[index + 1, 2] = Math.Round(item.d1, 2);
                    output[index + 1, 3] = Math.Round(item.r1, 2);
                    output[index + 1, 4] = Math.Round(item.d2, 2);
                    output[index + 1, 5] = Math.Round(item.r2, 2);
                    output[index + 1, 6] = Math.Round(item.H, 4);
                }
                writeRange.FormulaR1C1 = output;

                countRange = GetColumnRange(cell.Worksheet, startRow, countCol, results.Count);
                volumeRange = GetColumnRange(cell.Worksheet, startRow, vCol, results.Count);
                d1Range = GetColumnRange(cell.Worksheet, startRow, d1Col, results.Count);
                r1Range = GetColumnRange(cell.Worksheet, startRow, r1Col, results.Count);
                d2Range = GetColumnRange(cell.Worksheet, startRow, d2Col, results.Count);
                r2Range = GetColumnRange(cell.Worksheet, startRow, r2Col, results.Count);
                hRange = GetColumnRange(cell.Worksheet, startRow, hCol, results.Count);
                ApplyNumberFormatLocal(countRange, numberFormat4);
                ApplyNumberFormatLocal(volumeRange, numberFormat2);
                ApplyNumberFormatLocal(d1Range, numberFormat2);
                ApplyNumberFormatLocal(r1Range, numberFormat2);
                ApplyNumberFormatLocal(d2Range, numberFormat2);
                ApplyNumberFormatLocal(r2Range, numberFormat2);
                ApplyNumberFormatLocal(hRange, numberFormatH);

                summaryRange = cell.Worksheet.Range[
                    cell.Worksheet.Cells[cell.Row, countCol],
                    cell.Worksheet.Cells[cell.Row, vCol]];
                summaryRange.Font.Bold = true;
                summaryRange.Interior.ColorIndex = 0;
                summaryRange.NumberFormatLocal = numberFormat2;
                Excel.Range countSummary = null;
                try
                {
                    countSummary = cell.Worksheet.Cells[cell.Row, countCol] as Excel.Range;
                    countSummary.NumberFormatLocal = numberFormat4;
                }
                finally
                {
                    ReleaseComObject(countSummary);
                }
            }
            finally
            {
                ReleaseComObject(hRange);
                ReleaseComObject(r2Range);
                ReleaseComObject(d2Range);
                ReleaseComObject(r1Range);
                ReleaseComObject(d1Range);
                ReleaseComObject(volumeRange);
                ReleaseComObject(countRange);
                ReleaseComObject(summaryRange);
                ReleaseComObject(writeRange);
            }
        }


        static bool IsSingleColumnSelection(Excel.Range range)
        {
            int col = range.Cells[1, 1].Column;
            foreach (Excel.Range cell in range)
            {
                if (cell.Column != col)
                    return false;
            }
            return true;
        }

        static string GetAddress(Excel.Range rng) => rng.get_Address(false, false);

        static Excel.Range GetRangeFromAddress(Excel.Worksheet worksheet, string address)
        {
            if (worksheet == null || string.IsNullOrWhiteSpace(address))
                return null;

            try
            {
                return worksheet.Range[address.Trim()] as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        static void SetCellValue(Excel.Range cell, object value, string numberFormat = null, bool bold = false, int? fontColor = null)
        {
            if (value != null && value.ToString() != "Không tạo được")
                cell.Value2 = value;

            if (!string.IsNullOrEmpty(numberFormat))
                cell.NumberFormatLocal = numberFormat;

            cell.Font.Bold = bold;

            if (fontColor.HasValue)
                cell.Interior.Color = fontColor.Value;
            else
                cell.Interior.ColorIndex = 0;
        }

        static double RandomInRange(double min, double max)
        {
            return min + threadRandom.Value.NextDouble() * (max - min);
        }

        static int GetOptimalParallelism(DaodatRandomSettings settings)
        {
            if (settings.MaxParallelWorkers > 0)
                return settings.MaxParallelWorkers;

            int cpuCount = Environment.ProcessorCount;
            if (cpuCount <= 2)
                return Math.Max(1, cpuCount);

            return Math.Max(1, cpuCount - 1);
        }

        static List<double> GenerateRandomParts(int N, double min, double max, double targetSum, int maxAttempts)
        {
            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                HashSet<double> parts = new HashSet<double>();
                for (int i = 0; i < N; i++)
                    parts.Add(Math.Round(RandomInRange(min, max), 2));

                if (parts.Count != N) continue;

                double sum = parts.Sum();
                double scale = targetSum / sum;
                var scaled = parts.Select(x => Math.Round(x * scale, 3)).ToList();

                if (scaled.All(x => x >= min && x <= max) && scaled.Distinct().Count() == N)
                    return scaled;
            }
            return null;
        }

        static bool GenerateParametersForVolume(double V, out double d1, out double r1, out double d2, out double r2, out double H, DaodatRandomSettings settings)
        {
            for (int j = 0; j < settings.MaxAttempts; j++)
            {
                d1 = Math.Round(RandomInRange(settings.D1Min, settings.D1Max), 2);
                r1 = Math.Round(RandomInRange(settings.R1Min, settings.R1Max), 2);
                d2 = Math.Round(RandomInRange(settings.D2Min, settings.D2Max), 2);
                r2 = Math.Round(RandomInRange(settings.R2Min, settings.R2Max), 2);

                double A1 = d1 * r1;
                double A2 = d2 * r2;
                double denominator = A1 + A2 + Math.Sqrt(A1 * A2);

                H = 3 * V / denominator;

                if (H >= settings.HMin && H <= settings.HMax)
                {
                    H = Math.Round(H, 4);
                    return true;
                }
            }

            d1 = r1 = d2 = r2 = H = 0;
            return false;
        }
    }
}

