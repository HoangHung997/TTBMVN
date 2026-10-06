using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class EstimateV2ReportSheet
    {
        internal EstimateV2ReportSheet(string codeName, string name, string area)
        { CodeName = codeName; Name = name; PrintArea = area; }
        public string CodeName { get; }
        public string Name { get; }
        public string PrintArea { get; }
    }

    public sealed class EstimateV2ReportCheck
    {
        internal EstimateV2ReportCheck(IEnumerable<string> errors, IEnumerable<string> warnings)
        { Errors = Array.AsReadOnly(errors.Distinct().ToArray()); Warnings = Array.AsReadOnly(warnings.Distinct().ToArray()); }
        public IReadOnlyList<string> Errors { get; }
        public IReadOnlyList<string> Warnings { get; }
        public bool CanExport => Errors.Count == 0;
    }

    public static class WorkbookEstimateV2ReportService
    {
        public static IReadOnlyList<EstimateV2ReportSheet> ListSheets(Excel.Workbook workbook)
        {
            var result = new List<EstimateV2ReportSheet>();
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    Excel.Worksheet sheet = null;
                    Excel.PageSetup setup = null;
                    try
                    {
                        sheet = sheets.Item[i] as Excel.Worksheet;
                        if (sheet == null || sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible) continue;
                        setup = sheet.PageSetup;
                        string key = string.IsNullOrWhiteSpace(sheet.CodeName) ? "name:" + sheet.Name : sheet.CodeName;
                        result.Add(new EstimateV2ReportSheet(key, sheet.Name, setup.PrintArea ?? ""));
                    }
                    finally { Release(setup); Release(sheet); }
                }
            }
            finally { Release(sheets); }
            return result.AsReadOnly();
        }

        public static EstimateV2ReportCheck Check(Excel.Workbook workbook, string[] codeNames)
        {
            var errors = new List<string>();
            var warnings = new List<string>();
            if (codeNames == null || codeNames.Length == 0) errors.Add("Chưa chọn sheet xuất in.");
            foreach (string code in (codeNames ?? new string[0]).Distinct())
            {
                Excel.Worksheet sheet = null;
                Excel.PageSetup setup = null;
                Excel.Range area = null;
                Excel.Range intersection = null;
                Excel.Range used = null;
                try
                {
                    sheet = Resolve(workbook, code);
                    if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                        throw new InvalidOperationException("Sheet đã bị ẩn.");
                    setup = sheet.PageSetup;
                    if (string.IsNullOrWhiteSpace(setup.PrintArea))
                        throw new InvalidOperationException("Chưa đặt Print Area trong Excel.");
                    area = sheet.Range[setup.PrintArea];
                    sheet.Calculate();
                    foreach (Excel.XlCellType type in new[] { Excel.XlCellType.xlCellTypeFormulas, Excel.XlCellType.xlCellTypeConstants })
                    {
                        Excel.Range bad = null;
                        try
                        {
                            bad = area.SpecialCells(type, Excel.XlSpecialCellsValue.xlErrors);
                            intersection = workbook.Application.Intersect(area, bad);
                            if (intersection != null) errors.Add(sheet.Name + ": lỗi Excel tại " + intersection.Address[false, false]);
                        }
                        catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x800A03EC)) { }
                        finally { Release(intersection); intersection = null; Release(bad); }
                    }
                    // Managed metadata headers must never appear in an exported report.
                    used = sheet.UsedRange;
                    if (ContainsPrintedMetadata(sheet, used, area))
                        errors.Add(sheet.Name + ": vùng in chứa cột kỹ thuật V2.");
                }
                catch (Exception ex) { errors.Add(code + ": " + ex.Message); }
                finally { Release(used); Release(area); Release(setup); Release(sheet); }
            }
            if (WorkbookEstimateV2RegistrationService.ListRegistered(workbook).Count > 0)
            {
                try
                {
                    var validation = WorkbookEstimateV2ValidationService.ScanReadOnly(workbook);
                    errors.AddRange(validation.PackageErrors);
                    foreach (var finding in validation.Findings)
                    {
                        string text = finding.Code + ": " + finding.Title + " " + finding.Detail;
                        if (finding.Severity == EstimateV2CostIssueSeverity.Error) errors.Add(text);
                        else if (finding.Severity == EstimateV2CostIssueSeverity.Warning) warnings.Add(text);
                    }
                }
                catch (Exception ex) { errors.Add("Không kiểm tra được hồ sơ V2: " + ex.Message); }
            }
            else warnings.Add("Workbook chưa đăng ký công tác V2; chỉ xuất các bảng Excel hiện hành, chưa xác nhận tính đúng pháp lý/dự toán.");
            return new EstimateV2ReportCheck(errors, warnings);
        }

        public static void ExportPdf(Excel.Workbook workbook, string[] codeNames, string filePath, bool acceptWarnings)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !string.Equals(Path.GetExtension(filePath), ".pdf", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Chọn đường dẫn PDF hợp lệ.");
            string fullPath = Path.GetFullPath(filePath);
            if (!Directory.Exists(Path.GetDirectoryName(fullPath))) throw new DirectoryNotFoundException(fullPath);
            // Export beside the destination, then atomically replace: a failed export keeps the previous PDF.
            string temporary = Path.Combine(Path.GetDirectoryName(fullPath), Guid.NewGuid().ToString("N") + ".pdf");
            try
            {
                WithReport(workbook, codeNames, acceptWarnings, report => report.ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF, temporary, Excel.XlFixedFormatQuality.xlQualityStandard,
                    true, false, Type.Missing, Type.Missing, false, Type.Missing));
                if (!File.Exists(temporary) || new FileInfo(temporary).Length == 0) throw new IOException("Excel không tạo được PDF.");
                if (File.Exists(fullPath)) File.Replace(temporary, fullPath, null);
                else File.Move(temporary, fullPath);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        public static void PrintPreview(Excel.Workbook workbook, string[] codeNames, bool acceptWarnings)
        { WithReport(workbook, codeNames, acceptWarnings, report => report.PrintPreview(false)); }

        private static bool ContainsPrintedMetadata(Excel.Worksheet sheet, Excel.Range used, Excel.Range area)
        {
            Excel.Range rows = null, columns = null;
            Excel.Application application = null;
            try
            {
                rows = used.Rows;
                columns = used.Columns;
                application = sheet.Application;
                int firstRow = used.Row, firstColumn = used.Column;
                int lastRow = firstRow + rows.Count - 1, lastColumn = firstColumn + columns.Count - 1;
                // Find skips hidden cells on some Excel versions. Value2 includes them and ignores saved Find options.
                for (int c = firstColumn; c <= lastColumn; c += 128)
                for (int r = firstRow; r <= lastRow; r += 128)
                {
                    Excel.Range block = null;
                    try
                    {
                        string address = ExcelAddIn1.Core.ExcelColumnAddress.ToLetters(c) + r + ":" +
                            ExcelAddIn1.Core.ExcelColumnAddress.ToLetters(Math.Min(c + 127, lastColumn)) + Math.Min(r + 127, lastRow);
                        block = sheet.Range[address];
                        object raw = block.Value2;
                        var values = raw as object[,];
                        int height = values?.GetLength(0) ?? 1, width = values?.GetLength(1) ?? 1;
                        for (int i = 1; i <= height; i++)
                        for (int j = 1; j <= width; j++)
                        {
                            string value = (values == null ? raw : values[i, j]) as string;
                            if (value == null || !value.StartsWith("__TTB_", StringComparison.OrdinalIgnoreCase)) continue;
                            Excel.Range column = null, overlap = null;
                            try
                            {
                                string letter = ExcelAddIn1.Core.ExcelColumnAddress.ToLetters(c + j - 1);
                                column = sheet.Range[letter + ":" + letter];
                                overlap = application.Intersect(area, column);
                                if (overlap != null) return true;
                            }
                            finally { Release(overlap); Release(column); }
                        }
                    }
                    finally { Release(block); }
                }
                return false;
            }
            finally { Release(application); Release(columns); Release(rows); }
        }

        private static void WithReport(Excel.Workbook workbook, string[] codeNames, bool acceptWarnings, Action<Excel.Workbook> action)
        {
            EstimateV2ReportCheck check = Check(workbook, codeNames);
            if (!check.CanExport) throw new InvalidOperationException(string.Join(Environment.NewLine, check.Errors));
            if (check.Warnings.Count > 0 && !acceptWarnings) throw new InvalidOperationException("Cần xác nhận cảnh báo trước khi xuất.");
            Excel.Application app = workbook.Application;
            Excel.Workbook report = null;
            Excel.Sheets sheets = null;
            Excel.Sheets selected = null;
            object originalSheet = app.ActiveSheet;
            Excel.Workbook originalBook = app.ActiveWorkbook;
            try
            {
                using (var context = new ExcelWriteContext(app))
                {
                    var names = new List<object>();
                    foreach (string code in codeNames.Distinct())
                    {
                        Excel.Worksheet sheet = Resolve(workbook, code);
                        try { names.Add(sheet.Name); } finally { Release(sheet); }
                    }
                    sheets = workbook.Worksheets;
                    selected = sheets.Item[names.ToArray()] as Excel.Sheets;
                    if (selected == null) throw new InvalidOperationException("Không chọn được các sheet in.");
                    selected.Copy(Type.Missing, Type.Missing);
                    report = app.ActiveWorkbook;
                    if (report == null || string.Equals(report.FullName, workbook.FullName, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Không tạo được workbook in tạm.");
                    // The source stays open so copied formulas referring to it remain resolvable.
                    app.ScreenUpdating = true;
                    action(report);
                }
            }
            finally
            {
                try { report?.Close(false); }
                finally
                {
                    originalBook?.Activate();
                    (originalSheet as Excel.Worksheet)?.Activate();
                    Release(originalSheet); Release(originalBook); Release(selected); Release(sheets); Release(report); Release(app);
                }
            }
        }

        public static Excel.Worksheet Resolve(Excel.Workbook workbook, string codeName)
        {
            if (string.IsNullOrWhiteSpace(codeName)) throw new ArgumentException("Thiếu định danh sheet.");
            Excel.Sheets sheets = workbook.Worksheets;
            try
            {
                for (int i = 1; i <= sheets.Count; i++)
                {
                    var sheet = sheets.Item[i] as Excel.Worksheet;
                    if (sheet != null && (codeName.StartsWith("name:", StringComparison.Ordinal)
                        ? string.Equals(sheet.Name, codeName.Substring(5), StringComparison.OrdinalIgnoreCase)
                        : string.Equals(sheet.CodeName, codeName, StringComparison.OrdinalIgnoreCase))) return sheet;
                    Release(sheet);
                }
            }
            finally { Release(sheets); }
            throw new InvalidOperationException("Sheet không còn tồn tại: " + codeName);
        }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    }
}
