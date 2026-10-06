using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class EstimateV2RegisteredSource
    {
        public EstimateV2RegisteredSource(
            string worksheetCodeName,
            string worksheetName,
            string sourceAddress,
            int headerRow,
            int firstDataRow,
            int lastDataRow,
            EstimateV2ColumnLayout columns)
        {
            WorksheetCodeName = (worksheetCodeName ?? string.Empty).Trim();
            WorksheetName = (worksheetName ?? string.Empty).Trim();
            SourceAddress = (sourceAddress ?? string.Empty).Trim();
            HeaderRow = headerRow;
            FirstDataRow = firstDataRow;
            LastDataRow = lastDataRow;
            Columns = columns ?? throw new ArgumentNullException(nameof(columns));
        }

        public string WorksheetCodeName { get; }
        public string WorksheetName { get; }
        public string SourceAddress { get; }
        public int HeaderRow { get; }
        public int FirstDataRow { get; }
        public int LastDataRow { get; }
        public EstimateV2ColumnLayout Columns { get; }
    }

    public sealed class EstimateV2RegistrationResult
    {
        public EstimateV2RegistrationResult(
            EstimateV2RegisteredSource source,
            EstimateV2ReconcileResult reconcile)
        {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Reconcile = reconcile ?? throw new ArgumentNullException(nameof(reconcile));
        }

        public EstimateV2RegisteredSource Source { get; }
        public EstimateV2ReconcileResult Reconcile { get; }
    }

    public static class WorkbookEstimateV2RegistrationService
    {
        private const string Prefix = "TTBMVN.EstimateV2.";
        private const string PropertyManaged = Prefix + "Managed";
        private const string PropertySourceAddress = Prefix + "SourceAddress";
        private const string PropertyHeaderRow = Prefix + "HeaderRow";
        private const string PropertyFirstDataRow = Prefix + "FirstDataRow";
        private const string PropertyLastDataRow = Prefix + "LastDataRow";
        private const string PropertyWorkCodeColumn = Prefix + "WorkCodeColumn";
        private const string PropertyNormColumn = Prefix + "NormColumn";
        private const string PropertyDescriptionColumn = Prefix + "DescriptionColumn";
        private const string PropertyUnitColumn = Prefix + "UnitColumn";
        private const string PropertyQuantityColumn = Prefix + "QuantityColumn";
        private const string PropertyTechnicalIdColumn = Prefix + "TechnicalIdColumn";
        private const string PropertyTechnicalNormColumn = Prefix + "TechnicalNormColumn";
        private const string PropertyTechnicalKindColumn = Prefix + "TechnicalKindColumn";
        private const string PropertyTechnicalHashColumn = Prefix + "TechnicalHashColumn";

        public static EstimateV2RegistrationResult RegisterSelectedRange(
            Excel.Workbook workbook,
            Excel.Range range)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (range == null)
                throw new ArgumentNullException(nameof(range));
            if (range.Areas.Count != 1)
                throw new ArgumentException("Chi duoc dang ky mot vung lien tuc.", nameof(range));
            if (range.Rows.Count < 2)
                throw new ArgumentException("Vung cong tac phai co dong tieu de va it nhat mot dong du lieu.", nameof(range));

            Excel.Worksheet worksheet = null;
            try
            {
                worksheet = range.Worksheet;
                EnsureBelongsToWorkbook(workbook, worksheet);

                int headerRow = range.Row;
                int firstDataRow = headerRow + 1;
                int lastDataRow = range.Row + range.Rows.Count - 1;
                IReadOnlyDictionary<string, int> detected = DetectColumns(range);

                int workCodeColumn = RequiredDetected(detected, "workCode", "Mã công tác");
                int normColumn = RequiredDetected(detected, "norm", "Định mức");
                int descriptionColumn = RequiredDetected(detected, "description", "Mô tả công việc");
                int unitColumn = RequiredDetected(detected, "unit", "Đơn vị");
                int quantityColumn = RequiredDetected(detected, "quantity", "Khối lượng");

                int lastVisibleColumn = new[]
                {
                    workCodeColumn,
                    normColumn,
                    descriptionColumn,
                    unitColumn,
                    quantityColumn
                }.Max();
                EstimateV2ColumnLayout columns =
                    WorkbookEstimateV2StateService.EnsureTechnicalColumns(
                        worksheet,
                        headerRow,
                        workCodeColumn,
                        normColumn,
                        descriptionColumn,
                        unitColumn,
                        quantityColumn,
                        lastVisibleColumn + 1);

                string address = range.Address[
                    true,
                    true,
                    Excel.XlReferenceStyle.xlA1,
                    false,
                    Type.Missing];
                SaveSourceProperties(
                    worksheet,
                    address,
                    headerRow,
                    firstDataRow,
                    lastDataRow,
                    columns);

                EstimateV2ReconcileResult reconcile =
                    WorkbookEstimateV2StateService.Reconcile(
                        workbook,
                        worksheet,
                        firstDataRow,
                        lastDataRow,
                        columns);

                return new EstimateV2RegistrationResult(
                    new EstimateV2RegisteredSource(
                        worksheet.CodeName,
                        worksheet.Name,
                        address,
                        headerRow,
                        firstDataRow,
                        lastDataRow,
                        columns),
                    reconcile);
            }
            finally
            {
                Release(worksheet);
            }
        }

        public static IReadOnlyList<EstimateV2RegisteredSource> ListRegistered(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var result = new List<EstimateV2RegisteredSource>();
            Excel.Sheets sheets = null;
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
                        EstimateV2RegisteredSource source;
                        if (TryReadSource(worksheet, out source))
                            result.Add(source);
                    }
                    finally
                    {
                        Release(worksheet);
                    }
                }
                return new ReadOnlyCollection<EstimateV2RegisteredSource>(result);
            }
            finally
            {
                Release(sheets);
            }
        }

        public static EstimateV2ReconcileResult ReconcileAll(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            IReadOnlyList<EstimateV2RegisteredSource> sources = ListRegistered(workbook);
            if (sources.Count == 0)
            {
                return new EstimateV2ReconcileResult(
                    0, 0, 0, 0, 0, 0, false,
                    new[] { "Chua dang ky bang cong tac Du toan V2." });
            }

            int total = 0;
            int created = 0;
            int recovered = 0;
            int duplicates = 0;
            int restored = 0;
            int orphaned = 0;
            bool changed = false;
            var messages = new List<string>();

            foreach (EstimateV2RegisteredSource source in sources)
            {
                Excel.Worksheet worksheet = ResolveWorksheet(workbook, source);
                try
                {
                    EstimateV2ReconcileResult item =
                        WorkbookEstimateV2StateService.Reconcile(
                            workbook,
                            worksheet,
                            source.FirstDataRow,
                            source.LastDataRow,
                            source.Columns);
                    total += item.WorkItemCount;
                    created += item.CreatedCount;
                    recovered += item.RecoveredCount;
                    duplicates += item.DuplicateIdCount;
                    restored += item.RestoredNormDisplayCount;
                    orphaned += item.OrphanedCount;
                    changed |= item.StateChanged;
                    foreach (string message in item.Messages)
                        messages.Add(source.WorksheetName + ": " + message);
                }
                finally
                {
                    Release(worksheet);
                }
            }

            return new EstimateV2ReconcileResult(
                total,
                created,
                recovered,
                duplicates,
                restored,
                orphaned,
                changed,
                messages);
        }

        public static bool TryReadSource(
            Excel.Worksheet worksheet,
            out EstimateV2RegisteredSource source)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            Dictionary<string, string> values = ReadProperties(worksheet);
            string managed;
            if (!values.TryGetValue(PropertyManaged, out managed) ||
                !string.Equals(managed, "1", StringComparison.Ordinal))
            {
                source = null;
                return false;
            }

            try
            {
                var columns = new EstimateV2ColumnLayout(
                    Int(values, PropertyWorkCodeColumn),
                    Int(values, PropertyNormColumn),
                    Int(values, PropertyDescriptionColumn),
                    Int(values, PropertyUnitColumn),
                    Int(values, PropertyQuantityColumn),
                    Int(values, PropertyTechnicalIdColumn),
                    Int(values, PropertyTechnicalNormColumn),
                    Int(values, PropertyTechnicalKindColumn),
                    Int(values, PropertyTechnicalHashColumn));
                source = new EstimateV2RegisteredSource(
                    worksheet.CodeName,
                    worksheet.Name,
                    Value(values, PropertySourceAddress),
                    Int(values, PropertyHeaderRow),
                    Int(values, PropertyFirstDataRow),
                    Int(values, PropertyLastDataRow),
                    columns);
                return true;
            }
            catch (Exception ex) when (
                ex is KeyNotFoundException ||
                ex is FormatException ||
                ex is OverflowException ||
                ex is ArgumentException)
            {
                throw new InvalidOperationException(
                    "Metadata bang cong tac Du toan V2 tren sheet '" +
                    worksheet.Name + "' khong hop le.",
                    ex);
            }
        }

        private static IReadOnlyDictionary<string, int> DetectColumns(Excel.Range range)
        {
            object values = range.Value2;
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int index = 1; index <= range.Columns.Count; index++)
            {
                string header = NormalizeHeader(Text(values, 1, index));
                int column = range.Column + index - 1;

                Detect(result, "workCode", column, header,
                    "MA CONG TAC", "MA HIEU CONG TAC", "MA HIEU", "SO HIEU");
                Detect(result, "norm", column, header,
                    "DINH MUC", "MA DINH MUC", "SO HIEU DINH MUC");
                Detect(result, "description", column, header,
                    "MO TA CONG VIEC", "TEN CONG TAC", "NOI DUNG CONG VIEC", "NOI DUNG");
                Detect(result, "unit", column, header,
                    "DON VI", "DVT");
                Detect(result, "quantity", column, header,
                    "KHOI LUONG", "KL");
            }
            return new ReadOnlyDictionary<string, int>(result);
        }

        private static void Detect(
            IDictionary<string, int> result,
            string key,
            int column,
            string header,
            params string[] candidates)
        {
            if (result.ContainsKey(key) || header.Length == 0)
                return;
            foreach (string candidate in candidates)
            {
                if (header == candidate)
                {
                    result[key] = column;
                    return;
                }
            }
        }

        private static int RequiredDetected(
            IReadOnlyDictionary<string, int> detected,
            string key,
            string displayName)
        {
            int column;
            if (detected.TryGetValue(key, out column))
                return column;
            throw new InvalidOperationException(
                "Khong nhan dien duoc cot '" + displayName +
                "'. Hay chon vung co dong tieu de theo mau Du toan V2.");
        }

        private static string NormalizeHeader(string value)
        {
            string decomposed = (value ?? string.Empty).Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder();
            foreach (char ch in decomposed)
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (category != UnicodeCategory.NonSpacingMark)
                    builder.Append(char.ToUpperInvariant(ch));
            }
            string text = builder.ToString()
                .Replace('Đ', 'D')
                .Replace('đ', 'D');
            return string.Join(
                " ",
                text.Split(
                    new[] { ' ', '\t', '\r', '\n', '-', '_', '/', '(', ')' },
                    StringSplitOptions.RemoveEmptyEntries));
        }

        private static void SaveSourceProperties(
            Excel.Worksheet worksheet,
            string address,
            int headerRow,
            int firstDataRow,
            int lastDataRow,
            EstimateV2ColumnLayout columns)
        {
            SetProperty(worksheet, PropertyManaged, "1");
            SetProperty(worksheet, PropertySourceAddress, address);
            SetProperty(worksheet, PropertyHeaderRow, headerRow.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyFirstDataRow, firstDataRow.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyLastDataRow, lastDataRow.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyWorkCodeColumn, columns.WorkCodeColumn.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyNormColumn, columns.NormDisplayColumn.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyDescriptionColumn, columns.DescriptionColumn.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyUnitColumn, columns.UnitColumn.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyQuantityColumn, columns.QuantityColumn.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyTechnicalIdColumn, columns.TechnicalIdColumn.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyTechnicalNormColumn, columns.TechnicalNormColumn.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyTechnicalKindColumn, columns.TechnicalKindColumn.ToString(CultureInfo.InvariantCulture));
            SetProperty(worksheet, PropertyTechnicalHashColumn, columns.TechnicalFingerprintColumn.ToString(CultureInfo.InvariantCulture));
        }

        private static Dictionary<string, string> ReadProperties(Excel.Worksheet worksheet)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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
                        string name = Convert.ToString(property.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                        if (name.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase))
                        {
                            result[name] = Convert.ToString(
                                property.Value,
                                CultureInfo.InvariantCulture) ?? string.Empty;
                        }
                    }
                    finally
                    {
                        Release(property);
                    }
                }
                return result;
            }
            finally
            {
                Release(properties);
            }
        }

        private static void SetProperty(
            Excel.Worksheet worksheet,
            string name,
            string value)
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
                        if (!string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
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
                    first = properties.Add(name, value);
                else
                    first.Value = value;
            }
            finally
            {
                Release(first);
                Release(properties);
            }
        }

        private static Excel.Worksheet ResolveWorksheet(
            Excel.Workbook workbook,
            EstimateV2RegisteredSource source)
        {
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
                        if (string.Equals(
                            worksheet.CodeName,
                            source.WorksheetCodeName,
                            StringComparison.OrdinalIgnoreCase))
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
                throw new InvalidOperationException(
                    "Khong tim thay sheet cong tac Du toan V2 da dang ky.");
            }
            finally
            {
                Release(sheets);
            }
        }

        private static void EnsureBelongsToWorkbook(
            Excel.Workbook workbook,
            Excel.Worksheet worksheet)
        {
            Excel.Workbook parent = null;
            try
            {
                parent = worksheet.Parent as Excel.Workbook;
                if (parent == null || !string.Equals(
                    parent.FullName,
                    workbook.FullName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Vung cong tac phai thuoc workbook Du toan dang mo.");
                }
            }
            finally
            {
                Release(parent);
            }
        }

        private static string Value(
            IReadOnlyDictionary<string, string> values,
            string key)
        {
            string value;
            if (!values.TryGetValue(key, out value))
                throw new KeyNotFoundException(key);
            return value ?? string.Empty;
        }

        private static int Int(
            IReadOnlyDictionary<string, string> values,
            string key)
        {
            int value;
            if (!int.TryParse(
                Value(values, key),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out value))
            {
                throw new FormatException(key);
            }
            return value;
        }

        private static string Text(object matrix, int row, int column)
        {
            Array array = matrix as Array;
            object value = array == null
                ? (row == 1 && column == 1 ? matrix : null)
                : array.GetValue(
                    array.GetLowerBound(0) + row - 1,
                    array.GetLowerBound(1) + column - 1);
            return (Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty).Trim();
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
