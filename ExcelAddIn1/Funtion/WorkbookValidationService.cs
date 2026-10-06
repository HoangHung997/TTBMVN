using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookValidationService
    {
        private const decimal MoneyTolerance = 0.0001m;

        public static WorkbookValidationReport Scan(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            var issues = new List<WorkbookValidationIssue>();
            IReadOnlyList<WorksheetRoleAssignment> assignments = WorksheetRoleService.ReadAssignments(workbook);
            WorksheetRoleValidationResult roleValidation = WorksheetRoleValidator.Validate(assignments);
            foreach (WorksheetRoleValidationIssue issue in roleValidation.Issues)
            {
                issues.Add(new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.RoleMapping,
                    WorkbookValidationSeverity.Error,
                    issue.RoleId,
                    string.Empty,
                    string.Empty,
                    issue.RoleId,
                    issue.Message,
                    "Mo Thong tin chung va gan lai duy nhat mot sheet cho moi vai tro."));
            }

            ScanBrokenNames(workbook, issues);
            if (!roleValidation.IsValid)
                return new WorkbookValidationReport(issues);

            ProjectProfile project;
            PriceProfile profile;
            try
            {
                project = WorkbookProjectProfileService.LoadRequired(workbook);
                profile = WorkbookPriceProfileService.LoadRequired(workbook);
            }
            catch (Exception ex)
            {
                issues.Add(new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.InvalidProfile,
                    WorkbookValidationSeverity.Error,
                    WorksheetRoleCatalog.ToId(WorksheetRole.ResourcePrices),
                    string.Empty,
                    string.Empty,
                    "ProjectProfile/PriceProfile",
                    ex.Message,
                    "Mo Bang gia hoac Thong tin chung va luu lai ho so hop le."));
                return new WorkbookValidationReport(issues);
            }

            ScanAudit(workbook, assignments, project, profile, issues);
            ScanEstimate(workbook, issues);
            ScanCostSummary(workbook, issues);
            return new WorkbookValidationReport(issues);
        }

        public static bool NavigateTo(Excel.Workbook workbook, WorkbookValidationIssue issue)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (issue == null || issue.Address.Length == 0)
                return false;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            try
            {
                worksheet = ResolveByCodeName(workbook, issue.WorksheetCodeName);
                if (worksheet == null && issue.WorksheetRoleId.Length > 0 &&
                    WorksheetRoleCatalog.TryParse(issue.WorksheetRoleId, out WorksheetRole role))
                {
                    worksheet = WorksheetRoleService.ResolveRequired(workbook, role);
                }
                if (worksheet == null)
                    return false;
                range = worksheet.Range[issue.Address];
                worksheet.Activate();
                range.Select();
                return true;
            }
            finally
            {
                Release(range);
                Release(worksheet);
            }
        }

        private static void ScanAudit(
            Excel.Workbook workbook,
            IReadOnlyList<WorksheetRoleAssignment> assignments,
            ProjectProfile project,
            PriceProfile profile,
            ICollection<WorkbookValidationIssue> issues)
        {
            ResultAuditTrail trail;
            try
            {
                trail = WorkbookResultAuditService.Load(workbook);
            }
            catch (Exception ex)
            {
                issues.Add(new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.StaleAudit,
                    WorkbookValidationSeverity.Error,
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    "Result audit",
                    ex.Message,
                    "Ghi lai phu luc va tong hop kinh phi de tao audit moi."));
                return;
            }
            if (trail.Entries.Count == 0)
            {
                issues.Add(new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.StaleAudit,
                    WorkbookValidationSeverity.Warning,
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    "Result audit",
                    "Workbook chua co audit ket qua.",
                    "Ghi lai phu luc va tong hop kinh phi bang app."));
                return;
            }

            var assignmentByCodeName = assignments
                .Where(assignment => assignment.SheetKey.Length > 0)
                .GroupBy(assignment => assignment.SheetKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            var staleRoleScopes = new HashSet<string>(StringComparer.Ordinal);
            var staleIdentityScopes = new HashSet<string>(StringComparer.Ordinal);
            foreach (ResultAuditEntry entry in trail.Entries)
            {
                WorksheetRoleAssignment assignment;
                if (!assignmentByCodeName.TryGetValue(entry.WorksheetCodeName, out assignment) ||
                    !string.Equals(assignment.RoleId, entry.WorksheetRoleId, StringComparison.OrdinalIgnoreCase))
                {
                    staleRoleScopes.Add(entry.ScopeId + "|" + entry.WorksheetRoleId + "|" + entry.WorksheetCodeName);
                    continue;
                }
                if (!string.Equals(entry.PackageId, project.RegulationPackageId, StringComparison.Ordinal) ||
                    !string.Equals(entry.PackageVersion, project.RegulationPackageVersion, StringComparison.Ordinal) ||
                    !string.Equals(entry.PackageChecksum, project.RegulationPackageChecksum, StringComparison.OrdinalIgnoreCase) ||
                    (entry.PriceProfileId.Length > 0 &&
                        (!string.Equals(entry.PriceProfileId, profile.ProfileId, StringComparison.Ordinal) ||
                         !string.Equals(entry.PriceProfileVersion, profile.DataVersion, StringComparison.Ordinal) ||
                         !string.Equals(entry.PriceProfileChecksum, profile.Checksum, StringComparison.OrdinalIgnoreCase))))
                {
                    staleIdentityScopes.Add(entry.ScopeId + "|" + entry.WorksheetRoleId + "|" + entry.WorksheetCodeName);
                }
            }
            foreach (string value in staleRoleScopes)
            {
                string[] parts = value.Split('|');
                issues.Add(new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.StaleAudit,
                    WorkbookValidationSeverity.Error,
                    parts[1],
                    parts[2],
                    string.Empty,
                    parts[0],
                    "Audit khong con khop worksheet role/CodeName.",
                    "Gan lai role dung hoac ghi lai khoi ket qua."));
            }
            foreach (string value in staleIdentityScopes)
            {
                string[] parts = value.Split('|');
                issues.Add(new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.StaleAudit,
                    WorkbookValidationSeverity.Error,
                    parts[1],
                    parts[2],
                    string.Empty,
                    parts[0],
                    "Audit dung package/PriceProfile khac voi ho so dang pin.",
                    "Preview va ghi lai khoi ket qua bang package/PriceProfile hien tai."));
            }
        }

        private static void ScanEstimate(
            Excel.Workbook workbook,
            ICollection<WorkbookValidationIssue> issues)
        {
            Excel.Worksheet worksheet = null;
            try
            {
                WorkbookUnitRateContext context = WorkbookUnitRateService.LoadRequired(workbook);
                worksheet = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.EstimateAppendix);
                WorkbookEstimatePlan plan;
                if (!WorkbookResultAuditService.TryRestoreEstimatePlanForValidation(
                    workbook,
                    worksheet,
                    context,
                    out plan))
                {
                    plan = WorkbookEstimateAppendixService.ImportLegacyPlan(workbook, context);
                }
                IReadOnlyList<string> conditions = WorkbookResultAuditService.ReadEstimateConditions(
                    workbook,
                    worksheet.CodeName);
                WorkbookEstimatePreview preview = WorkbookEstimateAppendixService.Preview(
                    context,
                    plan,
                    conditions);
                foreach (WorkbookEstimatePreviewLine line in preview.Lines.Where(line => line.IsBlocking))
                {
                    WorkbookValidationIssueCode code = line.Error.IndexOf(
                        "Thieu gia",
                        StringComparison.OrdinalIgnoreCase) >= 0
                            ? WorkbookValidationIssueCode.MissingPrice
                            : WorkbookValidationIssueCode.StaleAudit;
                    issues.Add(new WorkbookValidationIssue(
                        code,
                        WorkbookValidationSeverity.Error,
                        WorksheetRoleCatalog.ToId(WorksheetRole.EstimateAppendix),
                        worksheet.CodeName,
                        "F" + line.Line.TargetRow.ToString(CultureInfo.InvariantCulture) + ":N" +
                            line.Line.TargetRow.ToString(CultureInfo.InvariantCulture),
                        line.Line.NormKey,
                        line.Error,
                        code == WorkbookValidationIssueCode.MissingPrice
                            ? "Bo sung gia dung ma/don vi trong Bang gia."
                            : "Kiem tra mapping dinh muc/binding va ghi lai phu luc."));
                }
                if (!preview.IsValid)
                    return;

                ScanFormulaMap(
                    worksheet,
                    WorksheetRole.EstimateAppendix,
                    WorkbookEstimateAppendixWriter.BuildExpectedFormulas(preview),
                    issues);
                CompareTotals(
                    worksheet,
                    WorksheetRole.EstimateAppendix,
                    new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
                    {
                        { "I27", preview.Result.MaterialAmountVnd },
                        { "J27", preview.Result.LaborAmountVnd },
                        { "K27", preview.Result.MachineAmountVnd }
                    },
                    issues);
            }
            catch (Exception ex)
            {
                issues.Add(ScanFailure(WorksheetRole.EstimateAppendix, worksheet?.CodeName, ex));
            }
            finally
            {
                Release(worksheet);
            }
        }

        private static void ScanCostSummary(
            Excel.Workbook workbook,
            ICollection<WorkbookValidationIssue> issues)
        {
            Excel.Worksheet summary = null;
            Excel.Worksheet estimate = null;
            try
            {
                WorkbookCostSummaryContext context = WorkbookCostSummaryService.Load(workbook);
                WorkbookCostSummaryPreview preview = WorkbookCostSummaryService.Preview(
                    context,
                    context.DefaultRequest);
                summary = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.CostSummary);
                estimate = WorksheetRoleService.ResolveRequired(workbook, WorksheetRole.EstimateAppendix);
                ScanFormulaMap(
                    summary,
                    WorksheetRole.CostSummary,
                    WorkbookCostSummaryWriter.BuildExpectedFormulas(estimate.Name, preview),
                    issues);
                CompareTotals(
                    summary,
                    WorksheetRole.CostSummary,
                    new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
                    {
                        { "E27", preview.Result.RoundedAfterTaxVnd }
                    },
                    issues);
            }
            catch (Exception ex)
            {
                issues.Add(ScanFailure(WorksheetRole.CostSummary, summary?.CodeName, ex));
            }
            finally
            {
                Release(estimate);
                Release(summary);
            }
        }

        private static void ScanFormulaMap(
            Excel.Worksheet worksheet,
            WorksheetRole role,
            IReadOnlyDictionary<string, string> expected,
            ICollection<WorkbookValidationIssue> issues)
        {
            if (expected.Count == 0)
                return;
            AddressCoordinate[] coordinates = expected.Keys.Select(ParseAddress).ToArray();
            int firstRow = coordinates.Min(item => item.Row);
            int lastRow = coordinates.Max(item => item.Row);
            int firstColumn = coordinates.Min(item => item.Column);
            int lastColumn = coordinates.Max(item => item.Column);
            Excel.Range range = null;
            try
            {
                range = worksheet.Range[
                    ExcelColumnAddress.ToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture),
                    ExcelColumnAddress.ToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture)];
                object formulas = range.Formula;
                foreach (KeyValuePair<string, string> item in expected)
                {
                    AddressCoordinate coordinate = ParseAddress(item.Key);
                    string actual = Convert.ToString(
                        MatrixValue(
                            formulas,
                            coordinate.Row - firstRow + 1,
                            coordinate.Column - firstColumn + 1),
                        CultureInfo.InvariantCulture) ?? string.Empty;
                    if (!FormulaEquals(item.Value, actual))
                    {
                        issues.Add(new WorkbookValidationIssue(
                            WorkbookValidationIssueCode.StaleFormula,
                            WorkbookValidationSeverity.Error,
                            WorksheetRoleCatalog.ToId(role),
                            worksheet.CodeName,
                            item.Key,
                            item.Key,
                            "Cong thuc khong khop. Expected=" + item.Value + "; Actual=" + actual + ".",
                            "Preview va ghi lai khoi ket qua bang app."));
                    }
                }
            }
            finally
            {
                Release(range);
            }
        }

        private static void CompareTotals(
            Excel.Worksheet worksheet,
            WorksheetRole role,
            IReadOnlyDictionary<string, decimal> expected,
            ICollection<WorkbookValidationIssue> issues)
        {
            if (expected.Count == 0)
                return;
            AddressCoordinate[] coordinates = expected.Keys.Select(ParseAddress).ToArray();
            int firstRow = coordinates.Min(item => item.Row);
            int lastRow = coordinates.Max(item => item.Row);
            int firstColumn = coordinates.Min(item => item.Column);
            int lastColumn = coordinates.Max(item => item.Column);
            Excel.Range range = null;
            try
            {
                range = worksheet.Range[
                    ExcelColumnAddress.ToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture),
                    ExcelColumnAddress.ToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture)];
                object values = range.Value2;
                foreach (KeyValuePair<string, decimal> item in expected)
                {
                    AddressCoordinate coordinate = ParseAddress(item.Key);
                    decimal actual = Convert.ToDecimal(
                        MatrixValue(
                            values,
                            coordinate.Row - firstRow + 1,
                            coordinate.Column - firstColumn + 1) ?? 0d,
                        CultureInfo.InvariantCulture);
                    if (Math.Abs(actual - item.Value) > MoneyTolerance)
                    {
                        issues.Add(new WorkbookValidationIssue(
                            WorkbookValidationIssueCode.IncorrectTotal,
                            WorkbookValidationSeverity.Error,
                            WorksheetRoleCatalog.ToId(role),
                            worksheet.CodeName,
                            item.Key,
                            item.Key,
                            "Tong sai. Expected=" + item.Value.ToString(CultureInfo.InvariantCulture) +
                                "; Actual=" + actual.ToString(CultureInfo.InvariantCulture) + ".",
                            "Kiem tra dau vao, recalculate va ghi lai khoi ket qua."));
                    }
                }
            }
            finally
            {
                Release(range);
            }
        }

        private static void ScanBrokenNames(
            Excel.Workbook workbook,
            ICollection<WorkbookValidationIssue> issues)
        {
            Excel.Names names = null;
            int inheritedCount = 0;
            var inheritedSamples = new List<string>();
            try
            {
                names = workbook.Names;
                for (int index = 1; index <= names.Count; index++)
                {
                    Excel.Name name = null;
                    try
                    {
                        name = names.Item(index);
                        string refersTo = Convert.ToString(name.RefersTo, CultureInfo.InvariantCulture) ?? string.Empty;
                        if (refersTo.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) < 0)
                            continue;
                        string nameText = Convert.ToString(name.Name, CultureInfo.InvariantCulture) ?? string.Empty;
                        bool applicationName = nameText.StartsWith("TTBMVN_", StringComparison.OrdinalIgnoreCase);
                        if (applicationName)
                        {
                            issues.Add(new WorkbookValidationIssue(
                                WorkbookValidationIssueCode.BrokenName,
                                WorkbookValidationSeverity.Error,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                nameText,
                                "Defined name tham chieu #REF!: " + refersTo + ".",
                                "Sua RefersTo hoac xoa defined name neu khong con su dung."));
                        }
                        else
                        {
                            inheritedCount++;
                            if (inheritedSamples.Count < 5)
                                inheritedSamples.Add(nameText);
                        }
                    }
                    catch (COMException ex)
                    {
                        RuntimeLogger.Log(ex, "Read workbook defined name");
                    }
                    finally
                    {
                        Release(name);
                    }
                }
            }
            finally
            {
                Release(names);
            }
            if (inheritedCount > 0)
            {
                issues.Add(new WorkbookValidationIssue(
                    WorkbookValidationIssueCode.BrokenName,
                    WorkbookValidationSeverity.Warning,
                    string.Empty,
                    string.Empty,
                    string.Empty,
                    inheritedCount.ToString(CultureInfo.InvariantCulture) + " defined name ke thua",
                    "Workbook co " + inheritedCount.ToString(CultureInfo.InvariantCulture) +
                        " defined name #REF! ke thua. Mau: " + string.Join(", ", inheritedSamples) + ".",
                    "Khong chan tinh toan; lap ke hoach don dep name khong con su dung.",
                    true));
            }
        }

        private static WorkbookValidationIssue ScanFailure(
            WorksheetRole role,
            string worksheetCodeName,
            Exception exception)
        {
            return new WorkbookValidationIssue(
                WorkbookValidationIssueCode.ScanFailure,
                WorkbookValidationSeverity.Error,
                WorksheetRoleCatalog.ToId(role),
                worksheetCodeName,
                string.Empty,
                WorksheetRoleCatalog.ToId(role),
                exception.Message,
                "Mo log runtime, sua loi dau vao va chay kiem tra lai.");
        }

        private static bool FormulaEquals(string expected, string actual)
        {
            return string.Equals(
                (expected ?? string.Empty).Replace(" ", string.Empty),
                (actual ?? string.Empty).Replace(" ", string.Empty),
                StringComparison.OrdinalIgnoreCase);
        }

        private static AddressCoordinate ParseAddress(string address)
        {
            Match match = Regex.Match(
                address ?? string.Empty,
                @"^(?<column>[A-Z]+)(?<row>[1-9][0-9]*)$",
                RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
            if (!match.Success)
                throw new ArgumentException("Dia chi cell khong hop le: " + address + ".", nameof(address));
            return new AddressCoordinate(
                ExcelColumnAddress.ToNumber(match.Groups["column"].Value),
                int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture));
        }

        private static object MatrixValue(object values, int row, int column)
        {
            Array array = values as Array;
            if (array == null)
                return row == 1 && column == 1 ? values : null;
            return array.GetValue(
                array.GetLowerBound(0) + row - 1,
                array.GetLowerBound(1) + column - 1);
        }

        private static string Address(ResultAuditEntry entry)
        {
            return ExcelColumnAddress.ToLetters(entry.FirstColumn) + entry.FirstRow.ToString(CultureInfo.InvariantCulture) +
                ":" + ExcelColumnAddress.ToLetters(entry.LastColumn) + entry.LastRow.ToString(CultureInfo.InvariantCulture);
        }

        private static Excel.Worksheet ResolveByCodeName(Excel.Workbook workbook, string codeName)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet worksheet = sheets.Item[index] as Excel.Worksheet;
                    if (worksheet == null)
                        continue;
                    if (string.Equals(worksheet.CodeName, codeName, StringComparison.OrdinalIgnoreCase))
                        return worksheet;
                    Release(worksheet);
                }
                return null;
            }
            finally
            {
                Release(sheets);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private sealed class AddressCoordinate
        {
            public AddressCoordinate(int column, int row)
            {
                Column = column;
                Row = row;
            }

            public int Column { get; }
            public int Row { get; }
        }
    }

    internal static class ResultAuditEntryValidationExtensions
    {
        internal static string AddressOrEmpty(this ResultAuditEntry entry)
        {
            if (entry == null)
                return string.Empty;
            return ExcelColumnAddress.ToLetters(entry.FirstColumn) + entry.FirstRow.ToString(CultureInfo.InvariantCulture) +
                ":" + ExcelColumnAddress.ToLetters(entry.LastColumn) + entry.LastRow.ToString(CultureInfo.InvariantCulture);
        }
    }
}
