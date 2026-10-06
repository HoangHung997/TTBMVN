using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    internal sealed class WorkbookPackageMigrationSnapshot
    {
        private readonly string projectPayload;
        private readonly string pricePayload;
        private readonly bool hadPricePayload;
        private readonly string auditPayload;
        private readonly bool hadAuditPayload;
        private readonly IReadOnlyList<RangeSnapshot> ranges;

        private WorkbookPackageMigrationSnapshot(
            string projectPayload,
            string pricePayload,
            bool hadPricePayload,
            string auditPayload,
            bool hadAuditPayload,
            IReadOnlyList<RangeSnapshot> ranges)
        {
            this.projectPayload = projectPayload;
            this.pricePayload = pricePayload;
            this.hadPricePayload = hadPricePayload;
            this.auditPayload = auditPayload;
            this.hadAuditPayload = hadAuditPayload;
            this.ranges = ranges;
            Fingerprint = ComputeFingerprint(
                projectPayload,
                pricePayload,
                hadPricePayload,
                auditPayload,
                hadAuditPayload,
                ranges);
        }

        internal string Fingerprint { get; }

        internal static WorkbookPackageMigrationSnapshot Capture(
            Excel.Workbook workbook,
            WorkbookEstimatePlan plan)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (plan == null)
                throw new ArgumentNullException(nameof(plan));

            string project;
            if (!WorkbookProjectProfileService.TryReadPayload(workbook, out project))
                throw new InvalidOperationException("Workbook chua co ProjectProfile de migration.");
            string price;
            bool hadPrice = WorkbookPriceProfileService.TryReadPayload(workbook, out price);
            string audit;
            bool hadAudit = WorkbookResultAuditService.TryReadPayload(workbook, out audit);

            Excel.Worksheet estimate = null;
            Excel.Worksheet summary = null;
            try
            {
                estimate = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.EstimateAppendix);
                summary = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.CostSummary);
                int firstLine = int.MaxValue;
                int lastLine = 0;
                foreach (WorkbookEstimateLine line in plan.Lines)
                {
                    firstLine = Math.Min(firstLine, line.TargetRow);
                    lastLine = Math.Max(lastLine, line.TargetRow);
                }
                int firstRow = Math.Max(1, firstLine - 1);
                int lastRow = Math.Max(lastLine, Math.Max(plan.AdjustmentRow, plan.GrandTotalRow));
                var captured = new List<RangeSnapshot>
                {
                    RangeSnapshot.Capture(
                        estimate,
                        "B" + firstRow.ToString(CultureInfo.InvariantCulture) + ":N" +
                            lastRow.ToString(CultureInfo.InvariantCulture)),
                    RangeSnapshot.Capture(summary, "D10:E27"),
                    RangeSnapshot.Capture(summary, "A28")
                };
                return new WorkbookPackageMigrationSnapshot(
                    project,
                    price,
                    hadPrice,
                    audit,
                    hadAudit,
                    captured.AsReadOnly());
            }
            finally
            {
                Release(summary);
                Release(estimate);
            }
        }

        internal void Restore(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            foreach (RangeSnapshot range in ranges)
                range.Restore(workbook);
            WorkbookProjectProfileService.RestorePayload(workbook, projectPayload);
            if (hadPricePayload)
                WorkbookPriceProfileService.RestorePayload(workbook, pricePayload);
            else
                WorkbookPriceProfileService.Clear(workbook);
            WorkbookResultAuditService.RestorePayload(
                workbook,
                hadAuditPayload ? auditPayload : null);
        }

        internal void VerifyRestored(Excel.Workbook workbook, WorkbookEstimatePlan plan)
        {
            WorkbookPackageMigrationSnapshot restored = Capture(workbook, plan);
            if (!string.Equals(Fingerprint, restored.Fingerprint, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Fingerprint workbook sau rollback khong khop snapshot.");
        }

        private static string ComputeFingerprint(
            string project,
            string price,
            bool hadPrice,
            string audit,
            bool hadAudit,
            IEnumerable<RangeSnapshot> capturedRanges)
        {
            var text = new StringBuilder();
            Append(text, "project", project);
            Append(text, "price-present", hadPrice ? "1" : "0");
            Append(text, "price", price);
            Append(text, "audit-present", hadAudit ? "1" : "0");
            Append(text, "audit", audit);
            foreach (RangeSnapshot range in capturedRanges)
            {
                Append(text, "range", range.WorksheetCodeName + "!" + range.Address);
                AppendValue(text, range.Formula);
            }
            return ProjectProfileSerializer.ComputeChecksum(text.ToString());
        }

        private static void Append(StringBuilder output, string key, string value)
        {
            output.Append(key).Append('=').Append(value ?? string.Empty).Append('\n');
        }

        private static void AppendValue(StringBuilder output, object value)
        {
            Array array = value as Array;
            if (array != null)
            {
                int rowLower = array.GetLowerBound(0);
                int rowUpper = array.GetUpperBound(0);
                int columnLower = array.GetLowerBound(1);
                int columnUpper = array.GetUpperBound(1);
                output.Append("array:").Append(rowUpper - rowLower + 1).Append('x')
                    .Append(columnUpper - columnLower + 1).Append('\n');
                for (int row = rowLower; row <= rowUpper; row++)
                {
                    for (int column = columnLower; column <= columnUpper; column++)
                        AppendScalar(output, array.GetValue(row, column));
                }
                return;
            }
            AppendScalar(output, value);
        }

        private static void AppendScalar(StringBuilder output, object value)
        {
            if (value == null)
            {
                output.Append("null\n");
                return;
            }
            IFormattable formattable = value as IFormattable;
            string text = formattable == null
                ? Convert.ToString(value, CultureInfo.InvariantCulture)
                : formattable.ToString(null, CultureInfo.InvariantCulture);
            output.Append(value.GetType().FullName).Append(':').Append(text).Append('\n');
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private sealed class RangeSnapshot
        {
            private RangeSnapshot(string worksheetCodeName, string address, object formula)
            {
                WorksheetCodeName = worksheetCodeName;
                Address = address;
                Formula = formula;
            }

            internal string WorksheetCodeName { get; }
            internal string Address { get; }
            internal object Formula { get; }

            internal static RangeSnapshot Capture(Excel.Worksheet worksheet, string address)
            {
                Excel.Range range = null;
                try
                {
                    range = worksheet.Range[address];
                    return new RangeSnapshot(worksheet.CodeName, address, range.Formula);
                }
                finally
                {
                    Release(range);
                }
            }

            internal void Restore(Excel.Workbook workbook)
            {
                Excel.Worksheet worksheet = null;
                Excel.Range range = null;
                try
                {
                    worksheet = ResolveWorksheet(workbook, WorksheetCodeName);
                    range = worksheet.Range[Address];
                    range.Formula = Formula;
                }
                finally
                {
                    Release(range);
                    Release(worksheet);
                }
            }

            private static Excel.Worksheet ResolveWorksheet(
                Excel.Workbook workbook,
                string codeName)
            {
                Excel.Sheets sheets = null;
                try
                {
                    sheets = workbook.Worksheets;
                    for (int index = 1; index <= sheets.Count; index++)
                    {
                        Excel.Worksheet candidate = sheets[index] as Excel.Worksheet;
                        if (candidate != null && string.Equals(
                            candidate.CodeName,
                            codeName,
                            StringComparison.Ordinal))
                        {
                            return candidate;
                        }
                        Release(candidate);
                    }
                }
                finally
                {
                    Release(sheets);
                }
                throw new InvalidOperationException("Khong tim thay worksheet CodeName " + codeName + ".");
            }
        }
    }
}
