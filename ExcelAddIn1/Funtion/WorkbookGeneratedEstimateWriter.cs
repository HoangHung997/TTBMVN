using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookGeneratedEstimateWriteResult
    {
        internal WorkbookGeneratedEstimateWriteResult(
            IEnumerable<string> worksheetNames,
            int rateCount,
            int linkedRowCount)
        {
            WorksheetNames = new ReadOnlyCollection<string>(worksheetNames.ToList());
            RateCount = rateCount;
            LinkedRowCount = linkedRowCount;
        }

        public IReadOnlyList<string> WorksheetNames { get; }
        public int RateCount { get; }
        public int LinkedRowCount { get; }
    }

    public static class WorkbookGeneratedEstimateWriter
    {
        private const string GeneratedRoleProperty = "TTBMVN.GeneratedEstimateRole";
        private const decimal VerificationTolerance = 0.01m;

        public static WorkbookGeneratedEstimateWriteResult Apply(
            Excel.Workbook workbook,
            WorkbookEstimateWorkspacePreview preview)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (preview == null)
                throw new ArgumentNullException(nameof(preview));
            if (!preview.IsValid)
                throw new InvalidOperationException("Preview con loi; khong the sinh bang don gia.");

            var createdSheets = new List<Excel.Worksheet>();
            var usedSheets = new List<Excel.Worksheet>();
            try
            {
                using (new ExcelWriteContext(workbook.Application))
                using (var transaction = new ExcelBatchWriteTransaction())
                {
                    var priceNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (MachineRateAudience audience in preview.Rates
                        .Select(rate => rate.Group.Identity.LaborAudience)
                        .Distinct()
                        .OrderBy(value => value))
                    {
                        Excel.Worksheet sheet = ResolveOrCreateGeneratedSheet(
                            workbook,
                            PriceRole(audience),
                            PriceSheetName(audience),
                            createdSheets);
                        usedSheets.Add(sheet);
                        WritePriceSheet(workbook, sheet, audience, preview, transaction, priceNames);
                    }

                    var rateNames = new Dictionary<string, RateFormulaNames>(StringComparer.Ordinal);
                    foreach (IGrouping<Tuple<MachineRateAudience, EstimateWorkEnvironment>, WorkbookEstimateRatePreview> group
                        in preview.Rates.GroupBy(rate => Tuple.Create(
                            rate.Group.Identity.LaborAudience,
                            rate.Group.Identity.Environment)))
                    {
                        Excel.Worksheet sheet = ResolveOrCreateGeneratedSheet(
                            workbook,
                            RateRole(group.Key.Item1, group.Key.Item2),
                            RateSheetName(group.Key.Item1, group.Key.Item2),
                            createdSheets);
                        usedSheets.Add(sheet);
                        WriteRateSheet(
                            workbook,
                            sheet,
                            group.Key.Item1,
                            group.Key.Item2,
                            group.OrderBy(rate => rate.Group.Identity.NormKey, StringComparer.Ordinal)
                                .ThenBy(rate => rate.Group.Identity.VariantCode, StringComparer.Ordinal)
                                .ToArray(),
                            transaction,
                            priceNames,
                            rateNames);
                    }

                    int linked = WriteAppendixLinks(workbook, preview, transaction, rateNames);
                    workbook.Application.Calculate();
                    VerifyRateTotals(workbook, preview, rateNames);
                    transaction.Commit();
                    return new WorkbookGeneratedEstimateWriteResult(
                        usedSheets.Select(sheet => sheet.Name),
                        preview.Rates.Count,
                        linked);
                }
            }
            catch
            {
                DeleteCreatedSheets(workbook, createdSheets);
                throw;
            }
            finally
            {
                foreach (Excel.Worksheet sheet in usedSheets)
                    Release(sheet);
                foreach (Excel.Worksheet sheet in createdSheets.Where(sheet => !usedSheets.Contains(sheet)))
                    Release(sheet);
            }
        }

        private static void WritePriceSheet(
            Excel.Workbook workbook,
            Excel.Worksheet sheet,
            MachineRateAudience audience,
            WorkbookEstimateWorkspacePreview preview,
            ExcelBatchWriteTransaction transaction,
            IDictionary<string, string> priceNames)
        {
            PriceProfile profile = preview.Context.PricePortfolio.FindRequired(audience);
            string[] priceCodes = preview.Rates
                .Where(rate => rate.Group.Identity.LaborAudience == audience)
                .SelectMany(rate => rate.Result.Resources)
                .Where(resource => !resource.IsPercentage && resource.PriceCode.Length > 0)
                .Select(resource => resource.PriceCode)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(code => code, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            int rows = Math.Max(4 + priceCodes.Length, ExistingRowCount(sheet));
            object[,] output = new object[rows, 8];
            output[0, 0] = "BẢNG GIÁ VẬT LIỆU - NHÂN CÔNG - MÁY";
            output[1, 0] = AudienceText(audience);
            output[1, 1] = profile.ProfileId + "@" + profile.DataVersion;
            string[] headers = { "Mã", "Loại", "Tên dữ liệu", "ĐVT", "Giá gốc", "Giá áp dụng", "Nguồn", "Lý do ghi đè" };
            for (int column = 0; column < headers.Length; column++)
                output[2, column] = headers[column];
            for (int index = 0; index < priceCodes.Length; index++)
            {
                PriceProfilePrice price = profile.FindRequired(priceCodes[index]);
                int row = index + 3;
                output[row, 0] = price.Entry.Code;
                output[row, 1] = PriceKindText(price.Entry.Kind);
                output[row, 2] = price.Entry.DisplayName;
                output[row, 3] = price.Entry.Unit;
                output[row, 4] = Number(price.Entry.BaseUnitPriceVnd);
                output[row, 5] = Number(price.AppliedUnitPriceVnd);
                output[row, 6] = price.Override?.SourceReference ?? price.Entry.SourceReference;
                output[row, 7] = price.Override?.Reason ?? string.Empty;
            }

            Excel.Range range = null;
            try
            {
                range = sheet.Range["A1", "H" + rows.ToString(CultureInfo.InvariantCulture)];
                transaction.WriteFormula(range, output);
                FormatPriceSheet(sheet, priceCodes.Length + 3);
            }
            finally
            {
                Release(range);
            }

            for (int index = 0; index < priceCodes.Length; index++)
            {
                string name = PriceName(audience, priceCodes[index]);
                int row = index + 4;
                SetWorkbookName(workbook, name, sheet, "F" + row.ToString(CultureInfo.InvariantCulture));
                priceNames[PriceDictionaryKey(audience, priceCodes[index])] = name;
            }
        }

        private static void WriteRateSheet(
            Excel.Workbook workbook,
            Excel.Worksheet sheet,
            MachineRateAudience audience,
            EstimateWorkEnvironment environment,
            IReadOnlyList<WorkbookEstimateRatePreview> rates,
            ExcelBatchWriteTransaction transaction,
            IReadOnlyDictionary<string, string> priceNames,
            IDictionary<string, RateFormulaNames> rateNames)
        {
            int requiredRows = 2;
            foreach (WorkbookEstimateRatePreview rate in rates)
                requiredRows += 8 + rate.Result.Resources.Count;
            int rows = Math.Max(requiredRows, ExistingRowCount(sheet));
            object[,] output = new object[rows, 8];
            output[0, 0] = "PHÂN TÍCH ĐƠN GIÁ " + EnvironmentText(environment).ToUpperInvariant();
            output[0, 1] = AudienceText(audience);
            int current = 2;
            foreach (WorkbookEstimateRatePreview rate in rates)
            {
                UnitRateCalculationResult result = rate.Result;
                output[current - 1, 0] = rate.Group.Identity.RateId;
                output[current - 1, 1] = result.NormKey + " / " + result.VariantCode;
                output[current - 1, 2] = result.NormTitle;
                output[current - 1, 6] = rate.Group.Rows.Count + " công tác sử dụng";
                output[current, 0] = "Loại";
                output[current, 1] = "Mã nguồn lực";
                output[current, 2] = "Tên nguồn lực";
                output[current, 3] = "ĐVT";
                output[current, 4] = "Hao phí";
                output[current, 5] = "Giá";
                output[current, 6] = "Thành tiền";
                output[current, 7] = "%";
                int firstResourceRow = current + 2;
                int lastResourceRow = firstResourceRow + result.Resources.Count - 1;
                for (int index = 0; index < result.Resources.Count; index++)
                {
                    UnitRateResourceAmount resource = result.Resources[index];
                    int arrayRow = current + 1 + index;
                    int excelRow = arrayRow + 1;
                    output[arrayRow, 0] = ResourceKindText(resource.Kind);
                    output[arrayRow, 1] = resource.ResourceCode;
                    output[arrayRow, 2] = resource.DisplayName;
                    output[arrayRow, 3] = resource.Unit;
                    output[arrayRow, 4] = Number(resource.Quantity);
                    output[arrayRow, 7] = resource.IsPercentage ? 1d : 0d;
                    if (resource.IsPercentage)
                    {
                        output[arrayRow, 5] = 0d;
                        output[arrayRow, 6] = "=SUMIFS($G$" + firstResourceRow + ":$G$" + lastResourceRow +
                            ",$A$" + firstResourceRow + ":$A$" + lastResourceRow +
                            ",\"VL\",$H$" + firstResourceRow + ":$H$" + lastResourceRow +
                            ",0)*E" + excelRow + "/100";
                    }
                    else
                    {
                        string priceName;
                        if (!priceNames.TryGetValue(
                            PriceDictionaryKey(audience, resource.PriceCode),
                            out priceName))
                        {
                            throw new InvalidOperationException("Khong co lien ket gia cho " + resource.PriceCode + ".");
                        }
                        output[arrayRow, 5] = "=" + priceName;
                        output[arrayRow, 6] = "=E" + excelRow + "*F" + excelRow;
                    }
                }
                int materialRow = lastResourceRow + 2;
                int laborRow = materialRow + 1;
                int machineRow = materialRow + 2;
                int totalRow = materialRow + 3;
                output[materialRow - 1, 5] = "Vật liệu";
                output[materialRow - 1, 6] = SumIfFormula(firstResourceRow, lastResourceRow, "VL");
                output[laborRow - 1, 5] = "Nhân công";
                output[laborRow - 1, 6] = SumIfFormula(firstResourceRow, lastResourceRow, "NC");
                output[machineRow - 1, 5] = "Máy";
                output[machineRow - 1, 6] = SumIfFormula(firstResourceRow, lastResourceRow, "M");
                output[totalRow - 1, 5] = "Đơn giá";
                output[totalRow - 1, 6] = "=SUM(G" + materialRow + ":G" + machineRow + ")";

                RateFormulaNames names = RateFormulaNames.Create(rate.Group.Identity.RateId);
                SetWorkbookName(workbook, names.Material, sheet, "G" + materialRow);
                SetWorkbookName(workbook, names.Labor, sheet, "G" + laborRow);
                SetWorkbookName(workbook, names.Machine, sheet, "G" + machineRow);
                SetWorkbookName(workbook, names.Total, sheet, "G" + totalRow);
                rateNames[rate.Group.Identity.CanonicalKey] = names;
                current = totalRow + 2;
            }

            Excel.Range range = null;
            try
            {
                range = sheet.Range["A1", "H" + rows.ToString(CultureInfo.InvariantCulture)];
                transaction.WriteFormula(range, output);
                FormatRateSheet(sheet, current - 1);
            }
            finally
            {
                Release(range);
            }
        }

        private static int WriteAppendixLinks(
            Excel.Workbook workbook,
            WorkbookEstimateWorkspacePreview preview,
            ExcelBatchWriteTransaction transaction,
            IReadOnlyDictionary<string, RateFormulaNames> rateNames)
        {
            int[] outputColumns = preview.Workspace.Source.Columns.AllColumns
                .Skip(5)
                .Where(column => column > 0)
                .Distinct()
                .OrderBy(column => column)
                .ToArray();
            if (outputColumns.Length == 0 || preview.Rates.Count == 0)
                return 0;
            Excel.Worksheet sheet = null;
            Excel.Range writeRange = null;
            try
            {
                sheet = WorkbookEstimateWorkspaceService.ResolveSourceWorksheet(
                    workbook,
                    preview.Workspace.Source);
                IReadOnlyDictionary<string, int> resolvedRows =
                    WorkbookEstimateWorkspaceService.ResolveBindingRows(workbook, preview.Workspace);
                int firstRow = preview.Workspace.Rows
                    .Where(row => !row.IsTextRow)
                    .Select(row => resolvedRows.ContainsKey(row.RowId) ? resolvedRows[row.RowId] : row.SourceRowHint)
                    .Min();
                int lastRow = preview.Workspace.Rows
                    .Where(row => !row.IsTextRow)
                    .Select(row => resolvedRows.ContainsKey(row.RowId) ? resolvedRows[row.RowId] : row.SourceRowHint)
                    .Max();
                int firstColumn = outputColumns.First();
                int lastColumn = outputColumns.Last();
                writeRange = sheet.Range[
                    ExcelColumnAddress.ToLetters(firstColumn) + firstRow,
                    ExcelColumnAddress.ToLetters(lastColumn) + lastRow];
                object[,] output = CloneMatrix(
                    writeRange.Formula,
                    lastRow - firstRow + 1,
                    lastColumn - firstColumn + 1);
                int linked = 0;
                foreach (WorkbookEstimateRatePreview rate in preview.Rates)
                {
                    RateFormulaNames names = rateNames[rate.Group.Identity.CanonicalKey];
                    foreach (EstimateWorkspaceRow row in rate.Group.Rows)
                    {
                        int targetRow;
                        if (!resolvedRows.TryGetValue(row.RowId, out targetRow))
                            targetRow = row.SourceRowHint;
                        Set(output, firstRow, firstColumn, targetRow,
                            preview.Workspace.Source.Columns.MaterialRateColumn, "=" + names.Material);
                        Set(output, firstRow, firstColumn, targetRow,
                            preview.Workspace.Source.Columns.LaborRateColumn, "=" + names.Labor);
                        Set(output, firstRow, firstColumn, targetRow,
                            preview.Workspace.Source.Columns.MachineRateColumn, "=" + names.Machine);
                        SetAmountFormula(output, preview.Workspace, firstRow, firstColumn, targetRow,
                            preview.Workspace.Source.Columns.MaterialAmountColumn,
                            preview.Workspace.Source.Columns.MaterialRateColumn,
                            names.Material);
                        SetAmountFormula(output, preview.Workspace, firstRow, firstColumn, targetRow,
                            preview.Workspace.Source.Columns.LaborAmountColumn,
                            preview.Workspace.Source.Columns.LaborRateColumn,
                            names.Labor);
                        SetAmountFormula(output, preview.Workspace, firstRow, firstColumn, targetRow,
                            preview.Workspace.Source.Columns.MachineAmountColumn,
                            preview.Workspace.Source.Columns.MachineRateColumn,
                            names.Machine);
                        linked++;
                    }
                }
                transaction.WriteFormula(writeRange, output);
                return linked;
            }
            finally
            {
                Release(writeRange);
                Release(sheet);
            }
        }

        private static void SetAmountFormula(
            object[,] output,
            EstimateWorkspace workspace,
            int firstRow,
            int firstColumn,
            int targetRow,
            int amountColumn,
            int rateColumn,
            string rateName)
        {
            if (amountColumn <= 0)
                return;
            string quantity = ExcelColumnAddress.ToLetters(workspace.Source.Columns.QuantityColumn) + targetRow;
            string rate = rateColumn > 0
                ? ExcelColumnAddress.ToLetters(rateColumn) + targetRow
                : rateName;
            Set(output, firstRow, firstColumn, targetRow, amountColumn, "=" + quantity + "*" + rate);
        }

        private static void VerifyRateTotals(
            Excel.Workbook workbook,
            WorkbookEstimateWorkspacePreview preview,
            IReadOnlyDictionary<string, RateFormulaNames> rateNames)
        {
            Excel.Names names = null;
            try
            {
                names = workbook.Names;
                foreach (WorkbookEstimateRatePreview rate in preview.Rates)
                {
                    RateFormulaNames formulaNames = rateNames[rate.Group.Identity.CanonicalKey];
                    VerifyName(names, formulaNames.Material, rate.Result.MaterialAmountVnd);
                    VerifyName(names, formulaNames.Labor, rate.Result.LaborAmountVnd);
                    VerifyName(names, formulaNames.Machine, rate.Result.MachineAmountVnd);
                    VerifyName(names, formulaNames.Total, rate.Result.TotalAmountVnd);
                }
            }
            finally
            {
                Release(names);
            }
        }

        private static void VerifyName(Excel.Names names, string name, decimal expected)
        {
            Excel.Name defined = null;
            Excel.Range range = null;
            try
            {
                defined = FindName(names, name);
                if (defined == null)
                    throw new InvalidOperationException("Khong tim thay name ket qua " + name + ".");
                range = defined.RefersToRange;
                decimal actual = Convert.ToDecimal(range.Value2 ?? 0d, CultureInfo.InvariantCulture);
                if (Math.Abs(actual - expected) > VerificationTolerance)
                {
                    throw new InvalidOperationException(
                        "Kiem tra " + name + " that bai: expected=" +
                        expected.ToString(CultureInfo.InvariantCulture) + ", actual=" +
                        actual.ToString(CultureInfo.InvariantCulture) + ".");
                }
            }
            finally
            {
                Release(range);
                Release(defined);
            }
        }

        private static Excel.Worksheet ResolveOrCreateGeneratedSheet(
            Excel.Workbook workbook,
            string role,
            string desiredName,
            ICollection<Excel.Worksheet> created)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index] as Excel.Worksheet;
                        if (sheet != null && string.Equals(
                            ReadGeneratedRole(sheet),
                            role,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Worksheet result = sheet;
                            sheet = null;
                            return result;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
                Excel.Worksheet added = sheets.Add(
                    Type.Missing,
                    sheets.Item[sheets.Count],
                    1,
                    Excel.XlSheetType.xlWorksheet) as Excel.Worksheet;
                added.Name = UniqueSheetName(workbook, desiredName);
                WriteGeneratedRole(added, role);
                created.Add(added);
                return added;
            }
            finally
            {
                Release(sheets);
            }
        }

        private static string ReadGeneratedRole(Excel.Worksheet worksheet)
        {
            Excel.CustomProperties properties = null;
            try
            {
                properties = worksheet.CustomProperties;
                for (int index = 1; index <= properties.Count; index++)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property = properties.Item[index];
                        if (string.Equals(property.Name, GeneratedRoleProperty, StringComparison.OrdinalIgnoreCase))
                            return Convert.ToString(property.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                    finally
                    {
                        Release(property);
                    }
                }
                return string.Empty;
            }
            finally
            {
                Release(properties);
            }
        }

        private static void WriteGeneratedRole(Excel.Worksheet worksheet, string role)
        {
            Excel.CustomProperties properties = null;
            Excel.CustomProperty property = null;
            try
            {
                properties = worksheet.CustomProperties;
                property = properties.Add(GeneratedRoleProperty, role);
            }
            finally
            {
                Release(property);
                Release(properties);
            }
        }

        private static void SetWorkbookName(
            Excel.Workbook workbook,
            string name,
            Excel.Worksheet worksheet,
            string address)
        {
            Excel.Names names = null;
            Excel.Name existing = null;
            Excel.Name created = null;
            Excel.Range cell = null;
            try
            {
                names = workbook.Names;
                existing = FindName(names, name);
                cell = worksheet.Range[address];
                string reference = "='" + worksheet.Name.Replace("'", "''") + "'!" +
                    cell.Address[true, true, Excel.XlReferenceStyle.xlA1, false, Type.Missing];
                if (existing != null)
                    existing.RefersTo = reference;
                else
                {
                    created = names.Add(
                        name,
                        reference,
                        false,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing);
                }
            }
            finally
            {
                Release(cell);
                Release(created);
                Release(existing);
                Release(names);
            }
        }

        private static Excel.Name FindName(Excel.Names names, string wanted)
        {
            try
            {
                return names.Item(wanted, Type.Missing, Type.Missing);
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string UniqueSheetName(Excel.Workbook workbook, string desired)
        {
            string baseName = desired.Length > 31 ? desired.Substring(0, 31) : desired;
            if (!SheetExists(workbook, baseName))
                return baseName;
            for (int index = 2; index < 1000; index++)
            {
                string suffix = " (" + index.ToString(CultureInfo.InvariantCulture) + ")";
                string prefix = baseName.Substring(0, Math.Min(baseName.Length, 31 - suffix.Length));
                string candidate = prefix + suffix;
                if (!SheetExists(workbook, candidate))
                    return candidate;
            }
            throw new InvalidOperationException("Khong tao duoc ten sheet cho " + desired + ".");
        }

        private static bool SheetExists(Excel.Workbook workbook, string name)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index] as Excel.Worksheet;
                        if (sheet != null && string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
                return false;
            }
            finally
            {
                Release(sheets);
            }
        }

        private static void DeleteCreatedSheets(Excel.Workbook workbook, IEnumerable<Excel.Worksheet> sheets)
        {
            bool alerts = workbook.Application.DisplayAlerts;
            try
            {
                workbook.Application.DisplayAlerts = false;
                foreach (Excel.Worksheet sheet in sheets.Reverse())
                {
                    try
                    {
                        sheet.Delete();
                    }
                    catch (Exception ex)
                    {
                        RuntimeLogger.Log(ex, "Rollback generated estimate sheet");
                    }
                }
            }
            finally
            {
                workbook.Application.DisplayAlerts = alerts;
            }
        }

        private static int ExistingRowCount(Excel.Worksheet sheet)
        {
            Excel.Range used = null;
            try
            {
                used = sheet.UsedRange;
                return Math.Max(1, used.Row + used.Rows.Count - 1);
            }
            finally
            {
                Release(used);
            }
        }

        private static void FormatPriceSheet(Excel.Worksheet sheet, int lastRow)
        {
            Excel.Range title = null;
            Excel.Range header = null;
            Excel.Range numbers = null;
            Excel.Range used = null;
            try
            {
                title = sheet.Range["A1", "H1"];
                title.Merge();
                title.Font.Bold = true;
                title.Font.Size = 14;
                title.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                header = sheet.Range["A3", "H3"];
                header.Font.Bold = true;
                header.Interior.Color = 0xD9EAD3;
                numbers = sheet.Range["E4", "F" + Math.Max(4, lastRow)];
                numbers.NumberFormat = "#,##0.00";
                used = sheet.Range["A1", "H" + Math.Max(3, lastRow)];
                used.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                used.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                sheet.Columns["A:H"].AutoFit();
                sheet.Columns["C:C"].ColumnWidth = 32;
                sheet.Columns["G:H"].ColumnWidth = 28;
            }
            finally
            {
                Release(used);
                Release(numbers);
                Release(header);
                Release(title);
            }
        }

        private static void FormatRateSheet(Excel.Worksheet sheet, int lastRow)
        {
            Excel.Range title = null;
            Excel.Range used = null;
            Excel.Range numbers = null;
            try
            {
                title = sheet.Range["A1", "H1"];
                title.Font.Bold = true;
                title.Font.Size = 14;
                numbers = sheet.Range["E1", "G" + Math.Max(1, lastRow)];
                numbers.NumberFormat = "#,##0.00";
                used = sheet.Range["A1", "H" + Math.Max(1, lastRow)];
                used.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                sheet.Columns["A:H"].AutoFit();
                sheet.Columns["C:C"].ColumnWidth = 42;
                sheet.Columns["H:H"].Hidden = true;
            }
            finally
            {
                Release(used);
                Release(numbers);
                Release(title);
            }
        }

        private static object[,] CloneMatrix(object source, int rows, int columns)
        {
            var output = new object[rows, columns];
            Array values = source as Array;
            if (values == null)
            {
                if (rows == 1 && columns == 1)
                    output[0, 0] = source;
                return output;
            }
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
            object[,] output,
            int firstRow,
            int firstColumn,
            int row,
            int column,
            object value)
        {
            if (column <= 0)
                return;
            output[row - firstRow, column - firstColumn] = value;
        }

        private static string SumIfFormula(int firstRow, int lastRow, string kind)
        {
            return "=SUMIF($A$" + firstRow + ":$A$" + lastRow + ",\"" + kind +
                "\",$G$" + firstRow + ":$G$" + lastRow + ")";
        }

        private static double Number(decimal value)
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        private static string PriceDictionaryKey(MachineRateAudience audience, string priceCode)
        {
            return ((int)audience).ToString(CultureInfo.InvariantCulture) + "|" + priceCode;
        }

        private static string PriceName(MachineRateAudience audience, string priceCode)
        {
            string prefix = audience == MachineRateAudience.StateBudgetSalary ? "HLNS" : "KHLNS";
            return "TTBMVN_GIA_" + prefix + "_" + Hash(priceCode).Substring(0, 14);
        }

        private static string Hash(string value)
        {
            using (SHA256 algorithm = SHA256.Create())
            {
                return string.Concat(algorithm.ComputeHash(Encoding.UTF8.GetBytes(value ?? string.Empty))
                    .Select(item => item.ToString("X2", CultureInfo.InvariantCulture)));
            }
        }

        private static string PriceRole(MachineRateAudience audience)
        {
            return "Price|" + ((int)audience).ToString(CultureInfo.InvariantCulture);
        }

        private static string RateRole(MachineRateAudience audience, EstimateWorkEnvironment environment)
        {
            return "Rate|" + ((int)audience).ToString(CultureInfo.InvariantCulture) + "|" +
                ((int)environment).ToString(CultureInfo.InvariantCulture);
        }

        private static string PriceSheetName(MachineRateAudience audience)
        {
            return audience == MachineRateAudience.StateBudgetSalary
                ? "VL-NC-M - HLNS"
                : "VL-NC-M - KHLNS";
        }

        private static string RateSheetName(
            MachineRateAudience audience,
            EstimateWorkEnvironment environment)
        {
            return "ĐG " + EnvironmentText(environment) + " - " +
                (audience == MachineRateAudience.StateBudgetSalary ? "HLNS" : "KHLNS");
        }

        private static string AudienceText(MachineRateAudience audience)
        {
            return audience == MachineRateAudience.StateBudgetSalary
                ? "Hưởng lương ngân sách"
                : "Không hưởng lương ngân sách";
        }

        private static string EnvironmentText(EstimateWorkEnvironment environment)
        {
            return environment == EstimateWorkEnvironment.Land ? "Cạn" : "Nước";
        }

        private static string ResourceKindText(NormResourceKind kind)
        {
            switch (kind)
            {
                case NormResourceKind.Material: return "VL";
                case NormResourceKind.Labor: return "NC";
                case NormResourceKind.Machine: return "M";
                default: throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        private static string PriceKindText(PriceResourceKind kind)
        {
            switch (kind)
            {
                case PriceResourceKind.Material: return "Vật liệu";
                case PriceResourceKind.Labor: return "Nhân công";
                case PriceResourceKind.FuelEnergy: return "Nhiên liệu";
                case PriceResourceKind.MachineOriginalPrice: return "Giá gốc máy";
                case PriceResourceKind.MachineShift: return "Ca máy";
                default: return kind.ToString();
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private sealed class RateFormulaNames
        {
            private RateFormulaNames(string material, string labor, string machine, string total)
            {
                Material = material;
                Labor = labor;
                Machine = machine;
                Total = total;
            }

            internal string Material { get; }
            internal string Labor { get; }
            internal string Machine { get; }
            internal string Total { get; }

            internal static RateFormulaNames Create(string rateId)
            {
                string token = (rateId ?? string.Empty).Replace("-", string.Empty);
                return new RateFormulaNames(
                    "TTBMVN_" + token + "_VL",
                    "TTBMVN_" + token + "_NC",
                    "TTBMVN_" + token + "_M",
                    "TTBMVN_" + token + "_TONG");
            }
        }
    }
}
