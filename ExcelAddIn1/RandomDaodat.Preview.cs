using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public sealed class DaodatRunPreview
    {
        public bool CanRun => Errors.Count == 0;
        public string Message { get; set; }
        public List<string> Errors { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();
    }

    public static partial class RandomDaodat
    {
        public static DaodatRunPreview BuildRunPreview(DaodatRunOptions runOptions)
        {
            if (runOptions == null)
                throw new ArgumentNullException(nameof(runOptions));

            return runOptions.HasTableData
                ? BuildTableDataPreview(runOptions)
                : BuildSelectionPreview(runOptions);
        }

        private static DaodatRunPreview BuildTableDataPreview(DaodatRunOptions runOptions)
        {
            var preview = new DaodatRunPreview();
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = xlApp.ActiveWorkbook;

            if (workbook == null)
                preview.Errors.Add("Chua co workbook Excel dang mo.");

            if (string.IsNullOrWhiteSpace(runOptions.TableDataAddress))
                preview.Errors.Add("Chua chon Table Data.");

            if (runOptions.OutputStartRow <= 0)
                preview.Errors.Add("Hang bat dau do ket qua phai lon hon 0.");

            ValidateSourceColumnsForPreview(runOptions, preview.Errors);
            ValidateColumnMappingsForPreview(runOptions.ColumnMappings, preview.Errors, preview.Warnings);

            Excel.Range dataRange = null;
            if (workbook != null && !string.IsNullOrWhiteSpace(runOptions.TableDataAddress))
            {
                dataRange = GetRangeFromWorkbookAddress(workbook, runOptions.TableDataAddress);
                if (dataRange == null)
                    preview.Errors.Add("Range Table Data khong hop le.");
            }

            string outputDescription = BuildOutputSheetDescription(workbook, runOptions, preview.Errors);
            Dictionary<string, string> outputColumns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                outputColumns = BuildOutputColumnMap(runOptions);
                AddGeneratedOutputWarnings(runOptions, outputColumns, preview.Warnings);
            }
            catch (Exception ex)
            {
                preview.Errors.Add(ex.Message);
            }

            int sourceRows = dataRange == null ? 0 : dataRange.Rows.Count;
            int validRows = 0;
            int totalOutputRows = 0;
            int total3mSignals = 0;
            int total5mSignals = 0;
            double total3mVolume = 0;
            double total5mVolume = 0;

            if (preview.Errors.Count == 0 && dataRange != null)
            {
                DaodatRandomSettings settings3m = GetSettings(false);
                DaodatRandomSettings settings5m = GetSettings(true);
                int firstRow = dataRange.Row;
                int rowCount = dataRange.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    int row = firstRow + i;
                    int th3 = ReadIntFromMappedColumn(dataRange.Worksheet, row, runOptions.SourceDataColumns, "TH3");
                    double v3 = ReadDoubleFromMappedColumn(dataRange.Worksheet, row, runOptions.SourceDataColumns, "V3");
                    int th5 = runOptions.ProjectIncludes5m ? ReadIntFromMappedColumn(dataRange.Worksheet, row, runOptions.SourceDataColumns, "TH5") : 0;
                    double v5 = runOptions.ProjectIncludes5m ? ReadDoubleFromMappedColumn(dataRange.Worksheet, row, runOptions.SourceDataColumns, "V5") : 0;

                    if (th3 <= 0 && th5 <= 0)
                        continue;

                    validRows++;
                    totalOutputRows += runOptions.ProjectIncludes5m ? Math.Max(th3, th5) + 1 : th3 + 1;

                    if (th3 > 0)
                    {
                        total3mSignals += th3;
                        total3mVolume += v3;
                        AddFeasibilityPreviewError(preview.Errors, $"Dong {row} V3", th3, v3, settings3m);
                    }

                    if (runOptions.ProjectIncludes5m && th5 > 0)
                    {
                        total5mSignals += th5;
                        total5mVolume += v5;
                        AddFeasibilityPreviewError(preview.Errors, $"Dong {row} V5", th5, v5, settings5m);
                    }
                }

                if (validRows == 0)
                    preview.Errors.Add("Khong co dong Table Data hop le de tinh.");
            }

            var lines = new List<string>
            {
                "Preview ho dao",
                $"Nguon: {runOptions.TableDataAddress}",
                $"Sheet dich: {outputDescription}",
                $"Hang bat dau: {runOptions.OutputStartRow.ToString(CultureInfo.CurrentCulture)}",
                $"Loai du an: {(runOptions.ProjectIncludes5m ? "Ca 3m va 5m" : "Chi 3m")}",
                $"So dong Table Data: {sourceRows.ToString(CultureInfo.CurrentCulture)}",
                $"So dong hop le: {validRows.ToString(CultureInfo.CurrentCulture)}",
                $"So dong se ghi: {(totalOutputRows + 1).ToString(CultureInfo.CurrentCulture)} (gom 1 dong header)",
                $"Tong TH3/V3: {total3mSignals.ToString(CultureInfo.CurrentCulture)} / {total3mVolume.ToString("0.###", CultureInfo.CurrentCulture)}",
                $"Tong TH5/V5: {total5mSignals.ToString(CultureInfo.CurrentCulture)} / {total5mVolume.ToString("0.###", CultureInfo.CurrentCulture)}",
                $"So cot output da map: {outputColumns.Count.ToString(CultureInfo.CurrentCulture)}",
                $"Link nguoc: {(runOptions.LinkBack ? "Co" : "Khong")}"
            };

            preview.Message = BuildPreviewMessage(lines, preview.Errors, preview.Warnings);
            return preview;
        }

        private static DaodatRunPreview BuildSelectionPreview(DaodatRunOptions runOptions)
        {
            var preview = new DaodatRunPreview();
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Excel.Worksheet sourceWorksheet = xlApp.ActiveSheet as Excel.Worksheet;
            Excel.Range selectedRange = xlApp.Selection as Excel.Range;

            if (sourceWorksheet == null)
                preview.Errors.Add("Chua co sheet Excel dang active.");

            if (selectedRange == null)
            {
                preview.Errors.Add("Vui long chon vung du lieu truoc khi chay.");
            }
            else if (!IsSingleColumnSelection(selectedRange))
            {
                preview.Errors.Add("Chi duoc chon cac o trong 1 cot.");
            }
            else if (selectedRange.Column <= 1)
            {
                preview.Errors.Add("Vung chon phai co cot ben trai de doc so luong N.");
            }

            if (runOptions.UseOutputStartCell && !runOptions.CreateNewSheet)
            {
                Excel.Range outputStartCell = GetRangeFromAddress(sourceWorksheet, runOptions.OutputStartAddress);
                if (outputStartCell == null)
                    preview.Errors.Add("Cell bat dau xuat khong hop le.");
            }

            int validCells = 0;
            int totalSignals = 0;
            double totalVolume = 0;
            DaodatRandomSettings settings = GetSettings(runOptions.Is5m);

            if (preview.Errors.Count == 0 && selectedRange != null)
            {
                int nCol = selectedRange.Column;
                int firstRow = selectedRange.Row;
                int lastRow = firstRow + selectedRange.Rows.Count - 1;
                for (int i = lastRow; i >= firstRow; i--)
                {
                    Excel.Range cell = sourceWorksheet.Cells[i, nCol] as Excel.Range;
                    if (cell?.Value2 == null)
                        continue;

                    if (!TryReadDouble(cell.Value2, out double targetSum) || targetSum == 0)
                        continue;

                    Excel.Range cellLeft = cell.Offset[0, -1] as Excel.Range;
                    if (cellLeft?.Value2 == null)
                        continue;

                    if (!TryReadInt(cellLeft.Value2, out int count) || count <= 0)
                        continue;

                    validCells++;
                    totalSignals += count;
                    totalVolume += targetSum;
                    AddFeasibilityPreviewError(preview.Errors, "o " + cell.Address[false, false], count, targetSum, settings);
                }

                if (validCells == 0)
                    preview.Errors.Add("Khong co o du lieu hop le de tinh.");
            }

            string outputMode = runOptions.CreateNewSheet ? "Tao sheet moi tu sheet dang chon" : "Ghi tren sheet hien tai";
            var lines = new List<string>
            {
                "Preview ho dao",
                $"Nguon: vung Excel dang chon",
                $"Sheet dich: {outputMode}",
                $"Loai ho dao: {(runOptions.Is5m ? "5m" : "3m")}",
                $"So o hop le: {validCells.ToString(CultureInfo.CurrentCulture)}",
                $"Tong so tin hieu: {totalSignals.ToString(CultureInfo.CurrentCulture)}",
                $"Tong V: {totalVolume.ToString("0.###", CultureInfo.CurrentCulture)}",
                $"So dong se chen: {(runOptions.InsertRows ? totalSignals : 0).ToString(CultureInfo.CurrentCulture)}",
                $"Link nguoc: {(runOptions.LinkBack ? "Co" : "Khong")}"
            };

            preview.Message = BuildPreviewMessage(lines, preview.Errors, preview.Warnings);
            return preview;
        }

        private static void ValidateSourceColumnsForPreview(DaodatRunOptions runOptions, List<string> errors)
        {
            if (runOptions.SourceDataColumns == null)
            {
                errors.Add("Chua khai bao cot data.");
                return;
            }

            RequireSourceColumn(runOptions, "TH3", errors);
            RequireSourceColumn(runOptions, "V3", errors);
            if (runOptions.ProjectIncludes5m)
            {
                RequireSourceColumn(runOptions, "TH5", errors);
                RequireSourceColumn(runOptions, "V5", errors);
            }

            foreach (KeyValuePair<string, string> item in runOptions.SourceDataColumns)
                ValidateColumnForPreview("Cot data " + item.Key, item.Value, errors);
        }

        private static void RequireSourceColumn(DaodatRunOptions runOptions, string key, List<string> errors)
        {
            if (!runOptions.SourceDataColumns.TryGetValue(key, out string column) || string.IsNullOrWhiteSpace(column))
                errors.Add("Chua khai bao cot data " + key + ".");
        }

        private static void ValidateColumnMappingsForPreview(
            List<DaodatColumnMapping> mappings,
            List<string> errors,
            List<string> warnings)
        {
            var usedOutputColumns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DaodatColumnMapping mapping in mappings ?? new List<DaodatColumnMapping>())
            {
                if (string.IsNullOrWhiteSpace(mapping.Key))
                    continue;

                if (!string.IsNullOrWhiteSpace(mapping.SourceColumn))
                {
                    if (IsSerialNumberSourceKey(mapping.SourceColumn))
                    {
                        warnings.Add(mapping.Key + " dung key_STT de danh so thu tu dong tong.");
                    }
                    else if (NormalizeKey(mapping.SourceColumn).StartsWith("KEY_", StringComparison.OrdinalIgnoreCase))
                    {
                        errors.Add("Key dac biet khong hop le cho " + mapping.Key + ": " + mapping.SourceColumn);
                    }
                    else
                    {
                        ValidateColumnForPreview("Cot trong Data cua " + mapping.Key, mapping.SourceColumn, errors);
                    }
                }

                if (!string.IsNullOrWhiteSpace(mapping.OutputColumn))
                {
                    try
                    {
                        string normalized = ExcelColumnAddress.Normalize(mapping.OutputColumn);
                        if (usedOutputColumns.TryGetValue(normalized, out string existingKey))
                            warnings.Add($"Cot output {normalized} dang duoc map cho ca {existingKey} va {mapping.Key}.");
                        else
                            usedOutputColumns[normalized] = mapping.Key;
                    }
                    catch (Exception ex)
                    {
                        errors.Add("Cot trong dao dap cua " + mapping.Key + " khong hop le. " + ex.Message);
                    }
                }
            }
        }

        private static void ValidateColumnForPreview(string label, string column, List<string> errors)
        {
            try
            {
                ExcelColumnAddress.Normalize(column);
            }
            catch (Exception ex)
            {
                errors.Add(label + " khong hop le. " + ex.Message);
            }
        }

        private static void AddGeneratedOutputWarnings(
            DaodatRunOptions runOptions,
            Dictionary<string, string> outputColumns,
            List<string> warnings)
        {
            if (outputColumns.Count == 0)
            {
                warnings.Add("Chua map cot output nao, ket qua se khong co du lieu do ra sheet dich.");
                return;
            }

            if (!HasAnyOutput(outputColumns, false))
                warnings.Add("Chua map cot output cho nhom 3m.");

            if (runOptions.ProjectIncludes5m && !HasAnyOutput(outputColumns, true))
                warnings.Add("Chua map cot output cho nhom 5m.");
        }

        private static bool HasAnyOutput(Dictionary<string, string> outputColumns, bool is5m)
        {
            string[] keys = is5m
                ? new[] { "TH5", "V5", "D1_5", "R1_5", "D2_5", "R2_5", "H_5" }
                : new[] { "TH3", "V3", "D1_3", "R1_3", "D2_3", "R2_3", "H_3" };

            return keys.Any(outputColumns.ContainsKey);
        }

        private static string BuildOutputSheetDescription(Excel.Workbook workbook, DaodatRunOptions runOptions, List<string> errors)
        {
            if (runOptions.CreateNewSheet)
                return "Tao sheet moi: " + (string.IsNullOrWhiteSpace(runOptions.NewSheetName) ? "HoDao" : runOptions.NewSheetName.Trim());

            if (string.IsNullOrWhiteSpace(runOptions.ExistingSheetName))
            {
                errors.Add("Chua chon sheet hien co.");
                return "Sheet hien co";
            }

            if (workbook != null && !SheetExists(workbook, runOptions.ExistingSheetName))
                errors.Add("Khong tim thay sheet dich: " + runOptions.ExistingSheetName);

            return "Sheet hien co: " + runOptions.ExistingSheetName;
        }

        private static void AddFeasibilityPreviewError(
            List<string> errors,
            string label,
            int count,
            double targetSum,
            DaodatRandomSettings settings)
        {
            if (targetSum <= 0)
            {
                errors.Add(label + ": V phai lon hon 0 khi so tin hieu > 0.");
                return;
            }

            DaodatFeasibilityResult feasibility = DaodatFeasibility.Analyze(count, targetSum, settings.VolumeMin, settings.VolumeMax);
            if (feasibility.CanGenerate)
                return;

            errors.Add(
                $"{label}: trung binh {feasibility.AverageVolume:0.###}; V co the tao {settings.VolumeMin:0.###}-{settings.VolumeMax:0.###}. {feasibility.Suggestion}");
        }

        private static string BuildPreviewMessage(List<string> lines, List<string> errors, List<string> warnings)
        {
            var builder = new StringBuilder();
            foreach (string line in lines)
                builder.AppendLine(line);

            AppendPreviewItems(builder, "Loi can sua", errors);
            AppendPreviewItems(builder, "Canh bao", warnings);

            if (errors.Count == 0)
                builder.AppendLine().Append("Trang thai: du dieu kien preview, chua ghi Excel.");

            return builder.ToString();
        }

        private static void AppendPreviewItems(StringBuilder builder, string title, List<string> items)
        {
            if (items == null || items.Count == 0)
                return;

            builder.AppendLine();
            builder.AppendLine(title + ":");
            int take = Math.Min(items.Count, 12);
            for (int i = 0; i < take; i++)
                builder.AppendLine("- " + items[i]);

            if (items.Count > take)
                builder.AppendLine("- ... con " + (items.Count - take).ToString(CultureInfo.CurrentCulture) + " dong nua.");
        }

        private static bool TryReadDouble(object value, out double number)
        {
            number = 0;
            if (value == null)
                return false;
            if (value is double doubleValue)
            {
                number = doubleValue;
                return true;
            }
            if (value is int intValue)
            {
                number = intValue;
                return true;
            }

            string text = value.ToString();
            return double.TryParse(text, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out number)
                || double.TryParse(text, System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number);
        }

        private static bool TryReadInt(object value, out int number)
        {
            number = 0;
            if (!TryReadDouble(value, out double doubleValue))
                return false;

            number = Convert.ToInt32(Math.Round(doubleValue, 0));
            return true;
        }
    }
}
