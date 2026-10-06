using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookEstimateWorkspaceService
    {
        public const string ManifestPropertyName = "TTBMVN.EstimateWorkspace.Manifest";
        private const string PartPropertyPrefix = "TTBMVN.EstimateWorkspace.Part.";
        private const string Label = "EstimateWorkspace";
        private const string RowBindingPrefix = "TTBMVN_PLDT_";
        private const string SourceIdProperty = "TTBMVN.EstimateSourceId";

        public static bool TryLoad(Excel.Workbook workbook, out EstimateWorkspace workspace)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            string payload;
            if (!WorkbookCustomPayloadStore.TryRead(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                Label,
                out payload))
            {
                workspace = null;
                return false;
            }
            workspace = EstimateWorkspaceSerializer.Deserialize(payload);
            return true;
        }

        public static EstimateWorkspace LoadRequired(Excel.Workbook workbook)
        {
            EstimateWorkspace workspace;
            if (!TryLoad(workbook, out workspace))
                throw new InvalidOperationException("Chua quet va luu vung Phu luc DT.");
            return workspace;
        }

        public static bool Save(Excel.Workbook workbook, EstimateWorkspace workspace)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            EstimateWorkspaceValidator.ValidateRequired(workspace);
            Excel.Worksheet worksheet = null;
            try
            {
                worksheet = ResolveSourceWorksheet(workbook, workspace.Source);
                SetSourceKey(worksheet, workspace.Source.WorksheetCodeName);
                EnsureRowBindings(workbook, worksheet, workspace);
                return WorkbookCustomPayloadStore.Save(
                    workbook,
                    ManifestPropertyName,
                    PartPropertyPrefix,
                    EstimateWorkspaceSerializer.Serialize(workspace),
                    Label);
            }
            finally
            {
                Release(worksheet);
            }
        }

        public static Excel.Worksheet ResolveSourceWorksheet(
            Excel.Workbook workbook,
            EstimateSourceBinding source)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            Excel.Sheets sheets = null;
            Excel.Worksheet nameFallback = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet worksheet = null;
                    try
                    {
                        worksheet = sheets.Item[index] as Excel.Worksheet;
                        if (worksheet == null)
                            continue;
                        if (string.Equals(worksheet.CodeName, source.WorksheetCodeName, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(ReadSourceKey(worksheet), source.WorksheetCodeName, StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Worksheet result = worksheet;
                            worksheet = null;
                            Release(nameFallback);
                            return result;
                        }
                        if (nameFallback == null && string.Equals(
                            worksheet.Name,
                            source.WorksheetName,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            nameFallback = worksheet;
                            worksheet = null;
                        }
                    }
                    finally
                    {
                        Release(worksheet);
                    }
                }
                if (nameFallback != null)
                    return nameFallback;
                throw new InvalidOperationException("Khong tim thay sheet Phu luc DT da quet.");
            }
            finally
            {
                Release(sheets);
            }
        }

        public static IReadOnlyDictionary<string, int> ResolveBindingRows(
            Excel.Workbook workbook,
            EstimateWorkspace workspace)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            var rows = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Excel.Names names = null;
            try
            {
                names = workbook.Names;
                foreach (EstimateWorkspaceRow row in workspace.Rows)
                {
                    Excel.Name name = null;
                    Excel.Range range = null;
                    try
                    {
                        name = FindName(names, row.BindingName);
                        if (name == null)
                            continue;
                        range = name.RefersToRange;
                        rows[row.RowId] = range.Row;
                    }
                    catch (COMException)
                    {
                    }
                    finally
                    {
                        Release(range);
                        Release(name);
                    }
                }
                return rows;
            }
            finally
            {
                Release(names);
            }
        }

        public static string CreateRowId()
        {
            return Guid.NewGuid().ToString("N");
        }

        public static string CreateBindingName(string rowId)
        {
            string id = (rowId ?? string.Empty).Trim().Replace("-", string.Empty);
            if (id.Length == 0)
                throw new ArgumentException("RowId khong duoc trong.", nameof(rowId));
            return RowBindingPrefix + id;
        }

        public static string GetOrCreateSourceKey(Excel.Worksheet worksheet)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            string codeName = (worksheet.CodeName ?? string.Empty).Trim();
            if (codeName.Length > 0)
                return codeName;
            string existing = ReadSourceKey(worksheet);
            if (existing.Length > 0)
                return existing;
            string created = "PLDT-" + Guid.NewGuid().ToString("N");
            SetSourceKey(worksheet, created);
            return created;
        }

        private static void EnsureRowBindings(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet,
            EstimateWorkspace workspace)
        {
            Excel.Names names = null;
            try
            {
                names = workbook.Names;
                foreach (EstimateWorkspaceRow row in workspace.Rows)
                {
                    Excel.Name existing = null;
                    Excel.Range cell = null;
                    Excel.Name created = null;
                    try
                    {
                        existing = FindName(names, row.BindingName);
                        if (existing != null)
                            continue;
                        int anchorColumn = workspace.Source.Columns.DescriptionColumn > 0
                            ? workspace.Source.Columns.DescriptionColumn
                            : workspace.Source.Columns.QuantityColumn;
                        cell = worksheet.Cells[row.SourceRowHint, anchorColumn] as Excel.Range;
                        string reference = "='" + worksheet.Name.Replace("'", "''") + "'!" +
                            cell.Address[true, true, Excel.XlReferenceStyle.xlA1, false, Type.Missing];
                        created = names.Add(
                            row.BindingName,
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
                    finally
                    {
                        Release(created);
                        Release(cell);
                        Release(existing);
                    }
                }
            }
            finally
            {
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

        private static string ReadSourceKey(Excel.Worksheet worksheet)
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
                        if (string.Equals(property.Name, SourceIdProperty, StringComparison.OrdinalIgnoreCase))
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

        private static void SetSourceKey(Excel.Worksheet worksheet, string key)
        {
            Excel.CustomProperties properties = null;
            Excel.CustomProperty first = null;
            try
            {
                properties = worksheet.CustomProperties;
                for (int index = properties.Count; index >= 1; index--)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property = properties.Item[index];
                        if (!string.Equals(property.Name, SourceIdProperty, StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (first == null)
                        {
                            first = property;
                            property = null;
                        }
                        else
                        {
                            property.Delete();
                        }
                    }
                    finally
                    {
                        Release(property);
                    }
                }
                if (first == null)
                    first = properties.Add(SourceIdProperty, key);
                else
                    first.Value = key;
            }
            finally
            {
                Release(first);
                Release(properties);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }

    public static class WorkbookEstimateRangeReader
    {
        public static EstimateWorkspace Read(
            Excel.Workbook workbook,
            Excel.Range selectedRange,
            EstimateColumnMap columns,
            bool firstRowIsHeader,
            EstimateWorkEnvironment defaultEnvironment,
            MachineRateAudience defaultAudience,
            EstimateWorkspace existingWorkspace)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (selectedRange == null)
                throw new ArgumentNullException(nameof(selectedRange));
            if (selectedRange.Areas.Count != 1)
                throw new ArgumentException("Chi duoc quet mot vung lien tuc.", nameof(selectedRange));

            Excel.Worksheet worksheet = null;
            try
            {
                worksheet = selectedRange.Worksheet;
                int firstRow = selectedRange.Row + (firstRowIsHeader ? 1 : 0);
                int lastRow = selectedRange.Row + selectedRange.Rows.Count - 1;
                int firstColumn = selectedRange.Column;
                int lastColumn = firstColumn + selectedRange.Columns.Count - 1;
                string sourceKey = WorkbookEstimateWorkspaceService.GetOrCreateSourceKey(worksheet);
                ValidateMappedColumns(columns, firstColumn, lastColumn);
                if (firstRow > lastRow)
                    throw new ArgumentException("Vung quet khong co dong du lieu.", nameof(selectedRange));

                object values = selectedRange.Value2;
                object formulas = selectedRange.FormulaLocal;
                Dictionary<int, EstimateWorkspaceRow> existingByRow = LoadExistingRows(
                    workbook,
                    existingWorkspace,
                    sourceKey,
                    firstRow,
                    lastRow);
                var rows = new List<EstimateWorkspaceRow>();
                for (int sourceRow = firstRow; sourceRow <= lastRow; sourceRow++)
                {
                    int rangeRow = sourceRow - selectedRange.Row + 1;
                    string code = Text(values, rangeRow, Relative(columns.CodeColumn, firstColumn));
                    string description = Text(values, rangeRow, Relative(columns.DescriptionColumn, firstColumn));
                    string unit = Text(values, rangeRow, Relative(columns.UnitColumn, firstColumn));
                    string quantityFormula = Formula(
                        formulas,
                        values,
                        rangeRow,
                        Relative(columns.QuantityColumn, firstColumn));
                    string acceptedFormula = Formula(
                        formulas,
                        values,
                        rangeRow,
                        Relative(columns.AcceptedQuantityColumn, firstColumn));
                    if (code.Length == 0 && description.Length == 0 && unit.Length == 0 &&
                        quantityFormula.Length == 0)
                    {
                        continue;
                    }

                    EstimateWorkspaceRow existing;
                    existingByRow.TryGetValue(sourceRow, out existing);
                    string rowId = existing?.RowId ?? WorkbookEstimateWorkspaceService.CreateRowId();
                    rows.Add(new EstimateWorkspaceRow(
                        rowId,
                        existing?.BindingName ?? WorkbookEstimateWorkspaceService.CreateBindingName(rowId),
                        sourceRow,
                        code,
                        description,
                        unit,
                        quantityFormula,
                        Decimal(
                            values,
                            rangeRow,
                            Relative(columns.QuantityColumn, firstColumn),
                            sourceRow,
                            "Khoi luong",
                            false),
                        acceptedFormula,
                        Decimal(
                            values,
                            rangeRow,
                            Relative(columns.AcceptedQuantityColumn, firstColumn),
                            sourceRow,
                            "Khoi luong nghiem thu",
                            true),
                        existing?.Environment ?? defaultEnvironment,
                        existing?.LaborAudience ?? defaultAudience,
                        existing?.NormKey ?? string.Empty,
                        existing?.VariantCode ?? string.Empty,
                        existing?.Conditions ?? Array.Empty<string>(),
                        existing?.Bindings ?? Array.Empty<UnitRateResourceBinding>()));
                }

                return new EstimateWorkspace(
                    new EstimateSourceBinding(
                        sourceKey,
                        worksheet.Name,
                        selectedRange.Address[true, true, Excel.XlReferenceStyle.xlA1, false, Type.Missing],
                        firstRow,
                        lastRow,
                        columns),
                    rows,
                    DateTime.UtcNow);
            }
            finally
            {
                Release(worksheet);
            }
        }

        private static Dictionary<int, EstimateWorkspaceRow> LoadExistingRows(
            Excel.Workbook workbook,
            EstimateWorkspace workspace,
            string worksheetCodeName,
            int firstRow,
            int lastRow)
        {
            var result = new Dictionary<int, EstimateWorkspaceRow>();
            if (workspace == null || !string.Equals(
                workspace.Source.WorksheetCodeName,
                worksheetCodeName,
                StringComparison.OrdinalIgnoreCase))
            {
                return result;
            }
            IReadOnlyDictionary<string, int> resolved =
                WorkbookEstimateWorkspaceService.ResolveBindingRows(workbook, workspace);
            foreach (EstimateWorkspaceRow row in workspace.Rows)
            {
                int currentRow;
                if (!resolved.TryGetValue(row.RowId, out currentRow))
                    currentRow = row.SourceRowHint;
                if (currentRow >= firstRow && currentRow <= lastRow && !result.ContainsKey(currentRow))
                    result.Add(currentRow, row);
            }
            return result;
        }

        private static void ValidateMappedColumns(EstimateColumnMap columns, int first, int last)
        {
            if (columns == null)
                throw new ArgumentNullException(nameof(columns));
            foreach (int column in columns.AllColumns.Where(value => value > 0))
            {
                if (column < first || column > last)
                    throw new ArgumentException("Cot " + ExcelColumnAddress.ToLetters(column) +
                        " nam ngoai vung quet.", nameof(columns));
            }
        }

        private static int Relative(int absoluteColumn, int firstColumn)
        {
            return absoluteColumn <= 0 ? 0 : absoluteColumn - firstColumn + 1;
        }

        private static string Text(object matrix, int row, int column)
        {
            if (column <= 0)
                return string.Empty;
            object value = MatrixValue(matrix, row, column);
            return (Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty).Trim();
        }

        private static string Formula(object formulas, object values, int row, int column)
        {
            if (column <= 0)
                return string.Empty;
            object formula = MatrixValue(formulas, row, column);
            string text = Convert.ToString(formula, CultureInfo.CurrentCulture) ?? string.Empty;
            if (text.StartsWith("=", StringComparison.Ordinal))
                return text;
            object value = MatrixValue(values, row, column);
            return value == null ? string.Empty : Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty;
        }

        private static decimal Decimal(
            object matrix,
            int row,
            int column,
            int sourceRow,
            string fieldName,
            bool optional)
        {
            if (column <= 0)
                return 0m;
            object value = MatrixValue(matrix, row, column);
            if (value == null || value is string && string.IsNullOrWhiteSpace((string)value))
                return 0m;
            if (IsExcelError(value))
            {
                if (optional)
                    return 0m;
                throw new ArgumentException(fieldName + " tai dong " + sourceRow +
                    " dang la loi Excel; hay sua cong thuc truoc khi quet.");
            }
            try
            {
                return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
            }
            catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
            {
                if (optional)
                    return 0m;
                throw new ArgumentException(fieldName + " tai dong " + sourceRow + " khong phai so.", ex);
            }
        }

        private static bool IsExcelError(object value)
        {
            if (value is ErrorWrapper)
                return true;
            if (!(value is int))
                return false;
            uint encoded = unchecked((uint)(int)value);
            return (encoded & 0xFFFF0000u) == 0x800A0000u;
        }

        private static object MatrixValue(object matrix, int row, int column)
        {
            Array array = matrix as Array;
            if (array == null)
                return row == 1 && column == 1 ? matrix : null;
            return array.GetValue(
                array.GetLowerBound(0) + row - 1,
                array.GetLowerBound(1) + column - 1);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
