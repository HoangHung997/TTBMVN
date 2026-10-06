using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public enum EstimateV2CostIssueSeverity
    {
        Info = 0,
        Warning = 1,
        Error = 2
    }

    public sealed class EstimateV2CostIssue
    {
        internal EstimateV2CostIssue(
            string code,
            string title,
            string detail,
            EstimateV2CostIssueSeverity severity)
            : this(
                code,
                title,
                detail,
                severity,
                string.Empty,
                string.Empty,
                false)
        {
        }

        internal EstimateV2CostIssue(
            string code,
            string title,
            string detail,
            EstimateV2CostIssueSeverity severity,
            string worksheetName,
            string address,
            bool recoverable)
        {
            Code = (code ?? string.Empty).Trim();
            Title = (title ?? string.Empty).Trim();
            Detail = (detail ?? string.Empty).Trim();
            Severity = severity;
            WorksheetName = (worksheetName ?? string.Empty).Trim();
            Address = (address ?? string.Empty).Trim();
            Recoverable = recoverable;
        }

        public string Code { get; }
        public string Title { get; }
        public string Detail { get; }
        public EstimateV2CostIssueSeverity Severity { get; }
        public string WorksheetName { get; }
        public string Address { get; }
        public bool Recoverable { get; }
    }

    public sealed class WorkbookEstimateV2CostPreview
    {
        internal WorkbookEstimateV2CostPreview(
            EstimateV2CostLinkPlan plan,
            int ratedWorkItemCount,
            int warningWorkItemCount,
            int formulaErrorWorkItemCount,
            bool thkpLinked,
            IEnumerable<string> packageErrors,
            IEnumerable<EstimateV2CostIssue> issues)
        {
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            RatedWorkItemCount = ratedWorkItemCount;
            WarningWorkItemCount = warningWorkItemCount;
            FormulaErrorWorkItemCount = formulaErrorWorkItemCount;
            ThkpLinked = thkpLinked;
            PackageErrors = new ReadOnlyCollection<string>(
                (packageErrors ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
            Issues = new ReadOnlyCollection<EstimateV2CostIssue>(
                (issues ?? Enumerable.Empty<EstimateV2CostIssue>())
                    .Where(item => item != null)
                    .ToList());
        }

        public EstimateV2CostLinkPlan Plan { get; }
        public int TotalWorkItemCount => Plan.TotalCount;
        public int RatedWorkItemCount { get; }
        public int WarningWorkItemCount { get; }
        public int FormulaErrorWorkItemCount { get; }
        public bool ThkpLinked { get; }
        public IReadOnlyList<string> PackageErrors { get; }
        public IReadOnlyList<EstimateV2CostIssue> Issues { get; }

        public bool CanUpdate =>
            TotalWorkItemCount > 0;
    }

    public static class WorkbookEstimateV2CostLinkService
    {
        public static WorkbookEstimateV2CostPreview BuildPreview(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            WorkbookEstimateV2ValidationReport validation =
                WorkbookEstimateV2ValidationService.Scan(workbook);

            var issues = validation.Findings
                .Select(item => new EstimateV2CostIssue(
                    item.Code,
                    item.Title,
                    item.Detail,
                    item.Severity,
                    item.WorksheetName,
                    item.Address,
                    item.Recoverable))
                .ToArray();

            return new WorkbookEstimateV2CostPreview(
                validation.Plan,
                validation.RatedWorkItemCount,
                validation.WarningWorkItemCount,
                validation.FormulaErrorWorkItemCount,
                validation.ThkpLinked,
                validation.PackageErrors,
                issues);
        }

        public static bool HasDirectCostLinks(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            return IsThkpDirectCostLinked(workbook);
        }

        internal static EstimateV2CostLinkPlan BuildLinkPlan(
            Excel.Workbook workbook,
            EstimateV2State state,
            ICollection<string> packageErrors)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (state == null)
                throw new ArgumentNullException(nameof(state));

            var rates = new List<EstimateV2RateItem>();
            foreach (EstimateV2RateEnvironment environment in new[]
            {
                EstimateV2RateEnvironment.Land,
                EstimateV2RateEnvironment.InlandWater,
                EstimateV2RateEnvironment.Sea
            })
            {
                try
                {
                    WorkbookEstimateV2RatePreview preview =
                        WorkbookEstimateV2RateService.BuildPreview(
                            workbook,
                            environment);
                    rates.AddRange(
                        preview.Items.Select(item => item.Rate));
                    if (packageErrors != null)
                    {
                        foreach (string error in
                            preview.MissingPackageBindings)
                        {
                            packageErrors.Add(error);
                        }
                    }
                }
                catch (Exception ex) when (
                    ex is ArgumentException ||
                    ex is InvalidOperationException ||
                    ex is KeyNotFoundException ||
                    ex is System.IO.IOException ||
                    ex is System.IO.InvalidDataException)
                {
                    if (packageErrors != null)
                        packageErrors.Add(ex.Message);
                }
            }

            EstimateV2RateItem[] uniqueRates = rates
                .GroupBy(
                    item => item.RateId,
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();

            return EstimateV2CostLinkPlan.Build(
                state,
                new EstimateV2RatePlan(uniqueRates));
        }

        private static int CountFormulaErrorWorkItems(
            Excel.Workbook workbook)
        {
            IReadOnlyList<EstimateV2RegisteredSource> sources =
                WorkbookEstimateV2RegistrationService.ListRegistered(
                    workbook);
            var erroredIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);

            foreach (EstimateV2RegisteredSource source in sources)
            {
                int firstOutput =
                    source.Columns.QuantityColumn + 1;
                int lastOutput =
                    source.Columns.QuantityColumn + 6;
                if (source.Columns.TechnicalIdColumn <= lastOutput)
                {
                    // Chua chay migration layout V2-401.
                    continue;
                }

                Excel.Worksheet sheet = null;
                try
                {
                    sheet = ResolveWorksheet(
                        workbook,
                        source);
                    for (int row = source.FirstDataRow;
                        row <= source.LastDataRow;
                        row++)
                    {
                        string id = ReadText(
                            sheet,
                            row,
                            source.Columns.TechnicalIdColumn);
                        if (!EstimateV2WorkItemState.IsValidId(id))
                            continue;

                        for (int column = firstOutput;
                            column <= lastOutput;
                            column++)
                        {
                            object value = ReadValue(
                                sheet,
                                row,
                                column);
                            if (IsExcelError(value))
                            {
                                erroredIds.Add(id);
                                break;
                            }
                        }
                    }
                }
                finally
                {
                    Release(sheet);
                }
            }

            return erroredIds.Count;
        }

        private static bool IsThkpDirectCostLinked(
            Excel.Workbook workbook)
        {
            Excel.Worksheet sheet = null;
            Excel.Range used = null;
            try
            {
                sheet =
                    WorkbookEstimateV2CompatibilityService
                        .ResolveOutputWorksheet(
                            workbook,
                            WorksheetRole.CostSummary,
                            "THKP-TC");
                if (sheet == null)
                    return false;

                string[] required =
                {
                    EstimateV2ExcelNames.EstimateTotal("VL"),
                    EstimateV2ExcelNames.EstimateTotal("NC"),
                    EstimateV2ExcelNames.EstimateTotal("M"),
                    EstimateV2ExcelNames.EstimateTotal("TOTAL")
                };

                Excel.Names names = null;
                try
                {
                    names = workbook.Names;
                    foreach (string name in required)
                    {
                        Excel.Name defined = null;
                        try
                        {
                            try
                            {
                                defined = names.Item(
                                    name,
                                    Type.Missing,
                                    Type.Missing);
                            }
                            catch (COMException)
                            {
                                return false;
                            }
                            if (defined == null)
                                return false;
                        }
                        finally
                        {
                            Release(defined);
                        }
                    }
                }
                finally
                {
                    Release(names);
                }

                var missingReferences =
                    new HashSet<string>(
                        required,
                        StringComparer.OrdinalIgnoreCase);
                used = sheet.UsedRange;
                object formulas = used.Formula;
                Array matrix = formulas as Array;
                if (matrix == null)
                {
                    RemoveFormulaReferences(
                        missingReferences,
                        Convert.ToString(formulas));
                }
                else
                {
                    foreach (object value in matrix)
                    {
                        RemoveFormulaReferences(
                            missingReferences,
                            Convert.ToString(value));
                        if (missingReferences.Count == 0)
                            break;
                    }
                }

                return missingReferences.Count == 0;
            }
            finally
            {
                Release(used);
                Release(sheet);
            }
        }

        private static void RemoveFormulaReferences(
            ISet<string> remaining,
            string formula)
        {
            string text = formula ?? string.Empty;
            if (!text.StartsWith(
                "=",
                StringComparison.Ordinal))
            {
                return;
            }

            foreach (string name in remaining.ToArray())
            {
                if (text.IndexOf(
                    name,
                    StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    remaining.Remove(name);
                }
            }
        }

        internal static Excel.Worksheet ResolveWorksheet(
            Excel.Workbook workbook,
            EstimateV2RegisteredSource source)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet nameFallback = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;

                        if (string.Equals(
                            sheet.CodeName,
                            source.WorksheetCodeName,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Worksheet result = sheet;
                            sheet = null;
                            Release(nameFallback);
                            return result;
                        }

                        if (nameFallback == null &&
                            string.Equals(
                                sheet.Name,
                                source.WorksheetName,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            nameFallback = sheet;
                            sheet = null;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }

                if (nameFallback != null)
                    return nameFallback;

                throw new InvalidOperationException(
                    "Khong tim thay bang cong tac da dang ky: " +
                    source.WorksheetName + ".");
            }
            finally
            {
                Release(sheets);
            }
        }

        internal static Excel.Worksheet FindWorksheet(
            Excel.Workbook workbook,
            string name)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet != null &&
                            string.Equals(
                                sheet.Name,
                                name,
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
                return null;
            }
            finally
            {
                Release(sheets);
            }
        }

        internal static string ReadText(
            Excel.Worksheet sheet,
            int row,
            int column)
        {
            object value = ReadValue(
                sheet,
                row,
                column);
            return (Convert.ToString(value) ??
                string.Empty).Trim();
        }

        internal static object ReadValue(
            Excel.Worksheet sheet,
            int row,
            int column)
        {
            Excel.Range cell = null;
            try
            {
                cell = sheet.Cells[row, column]
                    as Excel.Range;
                return cell?.Value2;
            }
            finally
            {
                Release(cell);
            }
        }

        internal static bool IsExcelError(object value)
        {
            if (value is ErrorWrapper)
                return true;
            if (!(value is int))
                return false;
            uint encoded = unchecked((uint)(int)value);
            return (encoded & 0xFFFF0000u) ==
                0x800A0000u;
        }

        internal static void Release(object value)
        {
            if (value != null &&
                Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
