using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class EstimateV2ValidationFinding
    {
        internal EstimateV2ValidationFinding(
            string code,
            string title,
            string detail,
            EstimateV2CostIssueSeverity severity,
            IEnumerable<string> workItemIds = null,
            string worksheetName = null,
            string address = null,
            bool recoverable = false)
        {
            Code = Clean(code);
            Title = Clean(title);
            Detail = Clean(detail);
            Severity = severity;
            WorkItemIds = new ReadOnlyCollection<string>(
                (workItemIds ?? Enumerable.Empty<string>())
                    .Select(Clean)
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
            WorksheetName = Clean(worksheetName);
            Address = Clean(address);
            Recoverable = recoverable;
        }

        public string Code { get; }
        public string Title { get; }
        public string Detail { get; }
        public EstimateV2CostIssueSeverity Severity { get; }
        public IReadOnlyList<string> WorkItemIds { get; }
        public string WorksheetName { get; }
        public string Address { get; }
        public bool Recoverable { get; }

        private static string Clean(string value)
        {
            return (value ?? string.Empty).Trim();
        }
    }

    public sealed class WorkbookEstimateV2ValidationReport
    {
        internal WorkbookEstimateV2ValidationReport(
            EstimateV2ReconcileResult reconcile,
            EstimateV2CostLinkPlan plan,
            int ratedWorkItemCount,
            int warningWorkItemCount,
            int formulaErrorWorkItemCount,
            int formulaOverwriteCount,
            int missingPriceWorkItemCount,
            bool thkpLinked,
            IEnumerable<string> packageErrors,
            IEnumerable<EstimateV2ValidationFinding> findings)
        {
            Reconcile = reconcile ?? throw new ArgumentNullException(nameof(reconcile));
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            RatedWorkItemCount = ratedWorkItemCount;
            WarningWorkItemCount = warningWorkItemCount;
            FormulaErrorWorkItemCount = formulaErrorWorkItemCount;
            FormulaOverwriteCount = formulaOverwriteCount;
            MissingPriceWorkItemCount = missingPriceWorkItemCount;
            ThkpLinked = thkpLinked;
            PackageErrors = new ReadOnlyCollection<string>(
                (packageErrors ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
            Findings = new ReadOnlyCollection<EstimateV2ValidationFinding>(
                (findings ?? Enumerable.Empty<EstimateV2ValidationFinding>())
                    .Where(item => item != null)
                    .OrderByDescending(item => item.Severity)
                    .ThenBy(item => item.Code, StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public EstimateV2ReconcileResult Reconcile { get; }
        public EstimateV2CostLinkPlan Plan { get; }
        public int TotalWorkItemCount => Plan.TotalCount;
        public int RatedWorkItemCount { get; }
        public int WarningWorkItemCount { get; }
        public int FormulaErrorWorkItemCount { get; }
        public int FormulaOverwriteCount { get; }
        public int MissingPriceWorkItemCount { get; }
        public bool ThkpLinked { get; }
        public IReadOnlyList<string> PackageErrors { get; }
        public IReadOnlyList<EstimateV2ValidationFinding> Findings { get; }
        public bool HasErrors => Findings.Any(item =>
            item.Severity == EstimateV2CostIssueSeverity.Error);
        public bool HasWarnings => Findings.Any(item =>
            item.Severity == EstimateV2CostIssueSeverity.Warning);
    }

    public sealed class WorkbookEstimateV2RepairResult
    {
        internal WorkbookEstimateV2RepairResult(
            WorkbookEstimateV2ValidationReport before,
            WorkbookEstimateV2ValidationReport after,
            int regeneratedEnvironmentCount,
            bool costLinksRebuilt,
            IEnumerable<string> errors)
        {
            Before = before ?? throw new ArgumentNullException(nameof(before));
            After = after ?? throw new ArgumentNullException(nameof(after));
            RegeneratedEnvironmentCount = regeneratedEnvironmentCount;
            CostLinksRebuilt = costLinksRebuilt;
            Errors = new ReadOnlyCollection<string>(
                (errors ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .ToList());
        }

        public WorkbookEstimateV2ValidationReport Before { get; }
        public WorkbookEstimateV2ValidationReport After { get; }
        public int RegeneratedEnvironmentCount { get; }
        public bool CostLinksRebuilt { get; }
        public IReadOnlyList<string> Errors { get; }
    }

    public static class WorkbookEstimateV2ValidationService
    {
        private static readonly string[] ManagedSheetNames =
        {
            "VL-NC-M",
            "DG Can",
            "DG Cạn",
            "DG Nuoc",
            "DG Nước",
            "DG Bien",
            "DG Biển",
            "THKP-TC"
        };

        public static WorkbookEstimateV2ValidationReport Scan(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            EstimateV2ReconcileResult reconcile;
            try
            {
                reconcile =
                    WorkbookEstimateV2RegistrationService
                        .ReconcileAll(workbook);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(
                    ex,
                    "V2 validation reconcile");
                reconcile = new EstimateV2ReconcileResult(
                    0, 0, 0, 0, 0, 0, false,
                    new[] { ex.Message });
            }

            EstimateV2State state;
            if (!WorkbookEstimateV2StateService.TryLoad(
                workbook,
                out state))
            {
                state = EstimateV2State.Empty(DateTime.UtcNow);
            }

            var findings = new List<EstimateV2ValidationFinding>();
            var packageErrors = new List<string>();
            var ratePreviews =
                new Dictionary<string, WorkbookEstimateV2RateItemPreview>(
                    StringComparer.OrdinalIgnoreCase);
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
                    foreach (WorkbookEstimateV2RateItemPreview item in
                        preview.Items)
                    {
                        ratePreviews[item.Rate.RateId] = item;
                        rates.Add(item.Rate);
                    }
                    packageErrors.AddRange(
                        preview.MissingPackageBindings);
                }
                catch (Exception ex) when (
                    ex is ArgumentException ||
                    ex is InvalidOperationException ||
                    ex is KeyNotFoundException ||
                    ex is System.IO.IOException ||
                    ex is System.IO.InvalidDataException)
                {
                    packageErrors.Add(ex.Message);
                }
            }

            EstimateV2RateItem[] uniqueRates = rates
                .GroupBy(
                    item => item.RateId,
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();

            EstimateV2CostLinkPlan plan =
                EstimateV2CostLinkPlan.Build(
                    state,
                    new EstimateV2RatePlan(uniqueRates));

            var generatedByRate =
                new Dictionary<string, bool>(
                    StringComparer.OrdinalIgnoreCase);
            foreach (string rateId in plan.Links
                .Where(item => item.IsReady)
                .Select(item => item.RateId)
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                generatedByRate[rateId] =
                    WorkbookEstimateV2RateService.IsGenerated(
                        workbook,
                        rateId);
            }

            int rated = plan.Links.Count(item =>
                item.IsReady &&
                generatedByRate.TryGetValue(
                    item.RateId,
                    out bool generated) &&
                generated);

            var warningWorkItems =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
            var formulaIssueWorkItems =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
            var overwrittenWorkItems =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
            var errorCells =
                new Dictionary<string, ErrorCell>(
                    StringComparer.OrdinalIgnoreCase);
            var overwrittenCells =
                new Dictionary<string, FormulaCell>(
                    StringComparer.OrdinalIgnoreCase);

            AddBindingFindings(
                plan,
                generatedByRate,
                ratePreviews,
                warningWorkItems,
                findings);

            int missingPriceWorkItems =
                plan.Links.Count(link =>
                {
                    WorkbookEstimateV2RateItemPreview item;
                    bool missing =
                        link.IsReady &&
                        ratePreviews.TryGetValue(
                            link.RateId,
                            out item) &&
                        item.HasMissingPrice;
                    return missing;
                });

            foreach (string packageError in packageErrors
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(3))
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "PACKAGE",
                    "Thiếu hoặc không khớp dữ liệu package",
                    packageError,
                    EstimateV2CostIssueSeverity.Warning));
            }

            if (reconcile.DuplicateIdCount > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "DUPLICATE_ID_RECOVERED",
                    "Đã tách " +
                        reconcile.DuplicateIdCount.ToString("N0") +
                        " WorkItemId bị trùng",
                    "Reconcile đã cấp ID mới cho dòng copy và giữ binding định mức.",
                    EstimateV2CostIssueSeverity.Info,
                    recoverable: true));
            }

            if (reconcile.RecoveredCount > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "IDENTITY_RECOVERED",
                    "Đã phục hồi " +
                        reconcile.RecoveredCount.ToString("N0") +
                        " dòng lệch WorkItemId",
                    "WorkItemId được phục hồi theo fingerprint sau chèn/xóa/sort.",
                    EstimateV2CostIssueSeverity.Info,
                    recoverable: true));
            }

            if (reconcile.RestoredNormDisplayCount > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "NORM_DISPLAY_RECOVERED",
                    "Đã phục hồi " +
                        reconcile.RestoredNormDisplayCount.ToString("N0") +
                        " ô hiển thị định mức",
                    "Binding thật trong Custom XML được dùng để khôi phục ô bị xóa/sửa.",
                    EstimateV2CostIssueSeverity.Info,
                    recoverable: true));
            }

            int orphaned = state.WorkItems.Count(item =>
                item != null && item.IsOrphaned);
            if (orphaned > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "ORPHAN",
                    "Có " +
                        orphaned.ToString("N0") +
                        " WorkItem orphan trong lịch sử",
                    "Các dòng đã bị xóa khỏi bảng được giữ trong state nhưng không tham gia tính toán.",
                    EstimateV2CostIssueSeverity.Info));
            }

            IReadOnlyList<EstimateV2RegisteredSource> sources =
                WorkbookEstimateV2RegistrationService.ListRegistered(
                    workbook);

            if (sources.Count == 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "NO_REGISTERED_SOURCE",
                    "Chưa đăng ký bảng công tác",
                    "Module vẫn mở bình thường; hãy đăng ký bảng công tác trước khi tổng hợp.",
                    EstimateV2CostIssueSeverity.Warning));
            }

            ScanPhysicalIdentity(
                workbook,
                state,
                sources,
                warningWorkItems,
                findings);

            ScanCostOutputFormulas(
                workbook,
                sources,
                plan,
                generatedByRate,
                warningWorkItems,
                formulaIssueWorkItems,
                overwrittenWorkItems,
                overwrittenCells,
                errorCells);

            ScanRateNamedCells(
                workbook,
                plan,
                generatedByRate,
                warningWorkItems,
                formulaIssueWorkItems,
                overwrittenWorkItems,
                overwrittenCells,
                errorCells,
                findings);

            ScanThkpFormulas(
                workbook,
                overwrittenCells,
                errorCells,
                findings);

            ScanManagedSheetErrors(
                workbook,
                errorCells);

            AddFormulaFindings(
                errorCells,
                overwrittenCells,
                warningWorkItems,
                formulaIssueWorkItems,
                overwrittenWorkItems,
                findings);

            bool thkpLinked =
                WorkbookEstimateV2CostLinkService
                    .HasDirectCostLinks(workbook) &&
                !overwrittenCells.Values.Any(item =>
                    string.Equals(
                        item.Code,
                        "THKP_FORMULA_OVERWRITTEN",
                        StringComparison.OrdinalIgnoreCase));
            if (thkpLinked)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "THKP_OK",
                    "Số liệu tổng hợp khớp",
                    "THKP-TC đang tham chiếu các workbook Name tổng VL/NC/M/T.",
                    EstimateV2CostIssueSeverity.Info,
                    worksheetName: "THKP-TC"));
            }
            else
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "THKP_LINK",
                    "THKP-TC chưa liên kết đầy đủ",
                    "Cập nhật THKP-TC để khôi phục bốn liên kết VL/NC/M/T.",
                    EstimateV2CostIssueSeverity.Warning,
                    worksheetName: "THKP-TC",
                    recoverable: true));
            }

            if (plan.UnboundCount == 0 &&
                plan.TotalCount > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "BINDING_OK",
                    "Tất cả công tác đã gắn định mức",
                    "Không còn WorkItem đang dùng thiếu binding định mức.",
                    EstimateV2CostIssueSeverity.Info));
            }

            if (rated == plan.TotalCount &&
                missingPriceWorkItems == 0 &&
                plan.TotalCount > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "RATE_OK",
                    "Đủ đơn giá cho tất cả công tác",
                    "Các RateId đang dùng đã sinh và không thiếu đầu vào giá.",
                    EstimateV2CostIssueSeverity.Info));
            }

            bool hasUnscopedFormulaProblem =
                errorCells.Values.Any(item =>
                    !EstimateV2WorkItemState.IsValidId(
                        item.WorkItemId)) ||
                overwrittenCells.Values.Any(item =>
                    !EstimateV2WorkItemState.IsValidId(
                        item.WorkItemId));
            int formulaErrorWorkItems =
                formulaIssueWorkItems.Count +
                (hasUnscopedFormulaProblem ? 1 : 0);
            if (errorCells.Count == 0 &&
                overwrittenCells.Count == 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "FORMULA_OK",
                    "Không phát hiện lỗi công thức",
                    "Không có #REF!, #VALUE!, #N/A hoặc công thức V2 bị ghi đè trong vùng quản lý.",
                    EstimateV2CostIssueSeverity.Info));
            }

            return new WorkbookEstimateV2ValidationReport(
                reconcile,
                plan,
                rated,
                warningWorkItems.Count,
                formulaErrorWorkItems,
                overwrittenCells.Count,
                missingPriceWorkItems,
                thkpLinked,
                packageErrors,
                findings);
        }

        public static WorkbookEstimateV2RepairResult RepairRecoverable(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            WorkbookEstimateV2ValidationReport before =
                Scan(workbook);
            var errors = new List<string>();
            int regenerated = 0;

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
                    if (preview.Items.Count == 0)
                        continue;

                    bool needsRepair =
                        before.Findings.Any(item =>
                            item.Recoverable &&
                            (item.Code == "RATE_FORMULA_OVERWRITTEN" ||
                             item.Code == "RATE_NAME_MISSING" ||
                             (item.Code == "MANAGED_FORMULA_ERROR" &&
                              (string.Equals(
                                   item.WorksheetName,
                                   "VL-NC-M",
                                   StringComparison.OrdinalIgnoreCase) ||
                               item.WorksheetName.StartsWith(
                                   "DG ",
                                   StringComparison.OrdinalIgnoreCase)))));

                    if (!needsRepair)
                        continue;

                    WorkbookEstimateV2RateSheetWriter.Apply(
                        workbook,
                        environment);
                    regenerated++;
                }
                catch (Exception ex)
                {
                    errors.Add(
                        environment + ": " + ex.Message);
                }
            }

            bool costLinksRebuilt = false;
            try
            {
                if (before.TotalWorkItemCount > 0 &&
                    before.Findings.Any(item =>
                        item.Recoverable &&
                        (item.Code == "COST_FORMULA_OVERWRITTEN" ||
                         item.Code == "THKP_FORMULA_OVERWRITTEN" ||
                         item.Code == "THKP_LINK" ||
                         (item.Code == "MANAGED_FORMULA_ERROR" &&
                          (item.WorksheetName.Length == 0 ||
                           string.Equals(
                               item.WorksheetName,
                               "THKP-TC",
                               StringComparison.OrdinalIgnoreCase) ||
                           !item.WorksheetName.StartsWith(
                               "DG ",
                               StringComparison.OrdinalIgnoreCase))))))
                {
                    WorkbookEstimateV2CostLinkWriter.Apply(
                        workbook);
                    costLinksRebuilt = true;
                }
            }
            catch (Exception ex)
            {
                errors.Add("Gia DT TC / THKP-TC: " + ex.Message);
            }

            WorkbookEstimateV2ValidationReport after =
                Scan(workbook);
            return new WorkbookEstimateV2RepairResult(
                before,
                after,
                regenerated,
                costLinksRebuilt,
                errors);
        }

        private static void AddBindingFindings(
            EstimateV2CostLinkPlan plan,
            IReadOnlyDictionary<string, bool> generatedByRate,
            IReadOnlyDictionary<string, WorkbookEstimateV2RateItemPreview> ratePreviews,
            ISet<string> warningWorkItems,
            ICollection<EstimateV2ValidationFinding> findings)
        {
            string[] unbound = plan.Links
                .Where(item =>
                    item.Status == EstimateV2CostLinkStatus.Unbound)
                .Select(item => item.WorkItemId)
                .ToArray();
            foreach (string id in unbound)
                warningWorkItems.Add(id);
            if (unbound.Length > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "MISSING_NORM",
                    "Còn " +
                        unbound.Length.ToString("N0") +
                        " công tác chưa gắn định mức",
                    "Gắn định mức cho các WorkItem này trước khi chốt hồ sơ.",
                    EstimateV2CostIssueSeverity.Warning,
                    unbound));
            }

            string[] unresolved = plan.Links
                .Where(item =>
                    item.Status ==
                        EstimateV2CostLinkStatus.RateNotResolved)
                .Select(item => item.WorkItemId)
                .ToArray();
            foreach (string id in unresolved)
                warningWorkItems.Add(id);
            if (unresolved.Length > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "RATE_NOT_RESOLVED",
                    unresolved.Length.ToString("N0") +
                        " công tác chưa resolve được RateId",
                    "Kiểm tra package, mã định mức và variant đã gắn.",
                    EstimateV2CostIssueSeverity.Warning,
                    unresolved));
            }

            string[] notGenerated = plan.Links
                .Where(item =>
                {
                    bool generated;
                    return item.IsReady &&
                        (!generatedByRate.TryGetValue(
                            item.RateId,
                            out generated) ||
                         !generated);
                })
                .Select(item => item.WorkItemId)
                .ToArray();
            foreach (string id in notGenerated)
                warningWorkItems.Add(id);
            if (notGenerated.Length > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "MISSING_RATE",
                    "Còn " +
                        notGenerated.Length.ToString("N0") +
                        " công tác chưa có đơn giá",
                    "Sinh/cập nhật DG Cạn, DG Nước hoặc DG Biển.",
                    EstimateV2CostIssueSeverity.Warning,
                    notGenerated,
                    recoverable: true));
            }

            string[] missingPrice = plan.Links
                .Where(item =>
                {
                    WorkbookEstimateV2RateItemPreview rate;
                    return item.IsReady &&
                        ratePreviews.TryGetValue(
                            item.RateId,
                            out rate) &&
                        rate.HasMissingPrice;
                })
                .Select(item => item.WorkItemId)
                .ToArray();
            foreach (string id in missingPrice)
                warningWorkItems.Add(id);
            if (missingPrice.Length > 0)
            {
                int missingRateCount = plan.Links
                    .Where(item => missingPrice.Contains(
                        item.WorkItemId,
                        StringComparer.OrdinalIgnoreCase))
                    .Select(item => item.RateId)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();
                findings.Add(new EstimateV2ValidationFinding(
                    "MISSING_PRICE",
                    "Thiếu giá đầu vào ở " +
                        missingRateCount.ToString("N0") +
                        " đơn giá",
                    missingPrice.Length.ToString("N0") +
                        " công tác bị ảnh hưởng. Cập nhật giá VL/NC/M trước khi xuất hồ sơ.",
                    EstimateV2CostIssueSeverity.Warning,
                    missingPrice,
                    "VL-NC-M"));
            }

            string[] condition = plan.Links
                .Where(item =>
                    item.IsReady &&
                    item.RequiresConditionReview)
                .Select(item => item.WorkItemId)
                .ToArray();
            foreach (string id in condition)
                warningWorkItems.Add(id);
            if (condition.Length > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "CONDITION_REVIEW",
                    condition.Length.ToString("N0") +
                        " công tác cần rà soát điều kiện định mức",
                    "Rate có adjustment/constraint vẫn đang dùng hao phí cơ sở.",
                    EstimateV2CostIssueSeverity.Warning,
                    condition));
            }
        }

        private static void ScanPhysicalIdentity(
            Excel.Workbook workbook,
            EstimateV2State state,
            IEnumerable<EstimateV2RegisteredSource> sources,
            ISet<string> warningWorkItems,
            ICollection<EstimateV2ValidationFinding> findings)
        {
            var seen =
                new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase);
            var duplicates =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
            var mismatches =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
            string firstSheet = string.Empty;
            string firstAddress = string.Empty;

            foreach (EstimateV2RegisteredSource source in sources)
            {
                Excel.Worksheet sheet = null;
                try
                {
                    sheet =
                        WorkbookEstimateV2CostLinkService
                            .ResolveWorksheet(
                                workbook,
                                source);
                    string sourceKey =
                        (sheet.CodeName ?? string.Empty).Trim();
                    if (sourceKey.Length == 0)
                        sourceKey =
                            (sheet.Name ?? string.Empty).Trim();

                    for (int row = source.FirstDataRow;
                        row <= source.LastDataRow;
                        row++)
                    {
                        string workCode = ReadText(
                            sheet,
                            row,
                            source.Columns.WorkCodeColumn);
                        string description = ReadText(
                            sheet,
                            row,
                            source.Columns.DescriptionColumn);
                        string unit = ReadText(
                            sheet,
                            row,
                            source.Columns.UnitColumn);
                        string quantity = ReadText(
                            sheet,
                            row,
                            source.Columns.QuantityColumn);
                        if (!IsWorkItemRow(
                            workCode,
                            description,
                            unit,
                            quantity))
                        {
                            continue;
                        }

                        string id = ReadText(
                            sheet,
                            row,
                            source.Columns.TechnicalIdColumn);
                        string address =
                            ExcelColumnAddress.ToLetters(
                                source.Columns.TechnicalIdColumn) +
                            row.ToString(
                                CultureInfo.InvariantCulture);

                        if (!EstimateV2WorkItemState.IsValidId(id))
                        {
                            mismatches.Add(
                                "ROW@" + sourceKey + ":" + row);
                            if (firstSheet.Length == 0)
                            {
                                firstSheet = sheet.Name;
                                firstAddress = address;
                            }
                            continue;
                        }

                        string previous;
                        if (seen.TryGetValue(id, out previous))
                        {
                            duplicates.Add(id);
                            warningWorkItems.Add(id);
                            if (firstSheet.Length == 0)
                            {
                                firstSheet = sheet.Name;
                                firstAddress = address;
                            }
                        }
                        else
                        {
                            seen[id] =
                                sourceKey + ":" + row;
                        }

                        EstimateV2WorkItemState item =
                            state.Find(id);
                        if (item == null ||
                            (item.SourceKey.Length > 0 &&
                             !string.Equals(
                                 item.SourceKey,
                                 sourceKey,
                                 StringComparison.OrdinalIgnoreCase)))
                        {
                            mismatches.Add(id);
                            warningWorkItems.Add(id);
                            if (firstSheet.Length == 0)
                            {
                                firstSheet = sheet.Name;
                                firstAddress = address;
                            }
                        }
                    }
                }
                finally
                {
                    Release(sheet);
                }
            }

            if (duplicates.Count > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "DUPLICATE_ID",
                    "Còn " +
                        duplicates.Count.ToString("N0") +
                        " WorkItemId trùng sau reconcile",
                    "Không được tính tiếp cho tới khi ID trùng được tách.",
                    EstimateV2CostIssueSeverity.Error,
                    duplicates,
                    firstSheet,
                    firstAddress,
                    true));
            }

            if (mismatches.Count > 0)
            {
                findings.Add(new EstimateV2ValidationFinding(
                    "IDENTITY_MISMATCH",
                    "Phát hiện " +
                        mismatches.Count.ToString("N0") +
                        " dòng lệch metadata/state",
                    "WorkItemId không hợp lệ, không tồn tại trong state hoặc thuộc source khác.",
                    EstimateV2CostIssueSeverity.Error,
                    mismatches.Where(
                        EstimateV2WorkItemState.IsValidId),
                    firstSheet,
                    firstAddress,
                    true));
            }
        }

        private static void ScanCostOutputFormulas(
            Excel.Workbook workbook,
            IEnumerable<EstimateV2RegisteredSource> sources,
            EstimateV2CostLinkPlan plan,
            IReadOnlyDictionary<string, bool> generatedByRate,
            ISet<string> warningWorkItems,
            ISet<string> formulaIssueWorkItems,
            ISet<string> overwrittenWorkItems,
            IDictionary<string, FormulaCell> overwrittenCells,
            IDictionary<string, ErrorCell> errorCells)
        {
            foreach (EstimateV2RegisteredSource source in sources)
            {
                int firstOutput =
                    source.Columns.QuantityColumn + 1;
                int lastOutput =
                    source.Columns.QuantityColumn + 6;
                if (source.Columns.TechnicalIdColumn <= lastOutput)
                    continue;

                Excel.Worksheet sheet = null;
                try
                {
                    sheet =
                        WorkbookEstimateV2CostLinkService
                            .ResolveWorksheet(
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

                        EstimateV2WorkItemCostLink link =
                            plan.Find(id);
                        bool generated;
                        bool expectFormula =
                            link != null &&
                            link.IsReady &&
                            generatedByRate.TryGetValue(
                                link.RateId,
                                out generated) &&
                            generated;
                        if (!expectFormula)
                            continue;

                        string[] expected =
                            EstimateV2ValidationRules
                                .ExpectedCostRowFormulas(
                                    source.Columns.QuantityColumn,
                                    row,
                                    link.RateId);

                        for (int index = 0;
                            index < 6;
                            index++)
                        {
                            int column = firstOutput + index;
                            Excel.Range cell = null;
                            try
                            {
                                cell = sheet.Cells[row, column]
                                    as Excel.Range;
                                string address =
                                    cell.Address[
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false,
                                        Type.Missing];
                                string key =
                                    CellKey(
                                        sheet,
                                        row,
                                        column);
                                string actual =
                                    Convert.ToString(
                                        cell.Formula,
                                        CultureInfo.InvariantCulture) ??
                                    string.Empty;

                                if (!EstimateV2ValidationRules
                                    .FormulaEquivalent(
                                        actual,
                                        expected[index]))
                                {
                                    overwrittenCells[key] =
                                        new FormulaCell(
                                            sheet.Name,
                                            address,
                                            id,
                                            actual,
                                            expected[index],
                                            "COST_FORMULA_OVERWRITTEN");
                                    overwrittenWorkItems.Add(id);
                                    formulaIssueWorkItems.Add(id);
                                    warningWorkItems.Add(id);
                                }

                                AddErrorIfAny(
                                    errorCells,
                                    sheet,
                                    cell,
                                    id);
                            }
                            finally
                            {
                                Release(cell);
                            }
                        }
                    }
                }
                finally
                {
                    Release(sheet);
                }
            }
        }

        private static void ScanRateNamedCells(
            Excel.Workbook workbook,
            EstimateV2CostLinkPlan plan,
            IReadOnlyDictionary<string, bool> generatedByRate,
            ISet<string> warningWorkItems,
            ISet<string> formulaIssueWorkItems,
            ISet<string> overwrittenWorkItems,
            IDictionary<string, FormulaCell> overwrittenCells,
            IDictionary<string, ErrorCell> errorCells,
            ICollection<EstimateV2ValidationFinding> findings)
        {
            foreach (string rateId in plan.Links
                .Where(item => item.IsReady)
                .Select(item => item.RateId)
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                bool generated;
                if (!generatedByRate.TryGetValue(
                    rateId,
                    out generated) ||
                    !generated)
                {
                    continue;
                }

                string[] affected = plan.Links
                    .Where(item =>
                        item.IsReady &&
                        string.Equals(
                            item.RateId,
                            rateId,
                            StringComparison.OrdinalIgnoreCase))
                    .Select(item => item.WorkItemId)
                    .ToArray();

                foreach (string component in new[]
                {
                    "VL", "NC", "M", "TOTAL"
                })
                {
                    string name =
                        EstimateV2ExcelNames.RateComponent(
                            rateId,
                            component);
                    Excel.Range cell = null;
                    try
                    {
                        if (!TryResolveNameRange(
                            workbook,
                            name,
                            out cell))
                        {
                            foreach (string id in affected)
                                warningWorkItems.Add(id);
                            findings.Add(
                                new EstimateV2ValidationFinding(
                                    "RATE_NAME_MISSING",
                                    "Thiếu workbook Name của đơn giá",
                                    rateId + " / " + component +
                                        " không còn Name để liên kết.",
                                    EstimateV2CostIssueSeverity.Error,
                                    affected,
                                    recoverable: true));
                            continue;
                        }

                        string formula =
                            Convert.ToString(
                                cell.Formula,
                                CultureInfo.InvariantCulture) ??
                            string.Empty;
                        string expectedFormula;
                        bool hasExactExpected =
                            TryBuildExpectedRateComponentFormula(
                                cell,
                                rateId,
                                component,
                                out expectedFormula);
                        bool overwritten =
                            !formula.StartsWith(
                                "=",
                                StringComparison.Ordinal) ||
                            (hasExactExpected &&
                             !EstimateV2ValidationRules
                                .FormulaEquivalent(
                                    formula,
                                    expectedFormula));
                        if (overwritten)
                        {
                            Excel.Worksheet sheet =
                                cell.Worksheet;
                            try
                            {
                                string address =
                                    cell.Address[
                                        false,
                                        false,
                                        Excel.XlReferenceStyle.xlA1,
                                        false,
                                        Type.Missing];
                                string key =
                                    CellKey(
                                        sheet,
                                        cell.Row,
                                        cell.Column);
                                overwrittenCells[key] =
                                    new FormulaCell(
                                        sheet.Name,
                                        address,
                                        string.Empty,
                                        formula,
                                        hasExactExpected
                                            ? expectedFormula
                                            : "formula",
                                        "RATE_FORMULA_OVERWRITTEN");
                                foreach (string id in affected)
                                {
                                    overwrittenWorkItems.Add(id);
                                    formulaIssueWorkItems.Add(id);
                                    warningWorkItems.Add(id);
                                }
                            }
                            finally
                            {
                                Release(sheet);
                            }
                        }

                        Excel.Worksheet errorSheet =
                            cell.Worksheet;
                        try
                        {
                            int errorCode;
                            if (TryGetExcelErrorCode(
                                cell.Value2,
                                out errorCode))
                            {
                                foreach (string id in affected)
                                {
                                    formulaIssueWorkItems.Add(id);
                                    warningWorkItems.Add(id);
                                }
                            }

                            AddErrorIfAny(
                                errorCells,
                                errorSheet,
                                cell,
                                affected.FirstOrDefault() ??
                                    string.Empty);
                        }
                        finally
                        {
                            Release(errorSheet);
                        }
                    }
                    finally
                    {
                        Release(cell);
                    }
                }
            }
        }

        private static bool TryBuildExpectedRateComponentFormula(
            Excel.Range cell,
            string rateId,
            string component,
            out string expected)
        {
            expected = string.Empty;
            if (cell == null)
                return false;

            string part =
                (component ?? string.Empty)
                    .Trim()
                    .ToUpperInvariant();
            if (part == "TOTAL")
            {
                if (cell.Row <= 1)
                    return false;
                int sourceRow = cell.Row - 1;
                expected =
                    "=SUM(F" +
                    sourceRow.ToString(
                        CultureInfo.InvariantCulture) +
                    ":H" +
                    sourceRow.ToString(
                        CultureInfo.InvariantCulture) +
                    ")";
                return true;
            }

            int amountColumn;
            switch (part)
            {
                case "VL":
                    amountColumn = 6;
                    break;
                case "NC":
                    amountColumn = 7;
                    break;
                case "M":
                    amountColumn = 8;
                    break;
                default:
                    return false;
            }

            Excel.Worksheet sheet = null;
            try
            {
                sheet = cell.Worksheet;
                int firstRow =
                    Math.Max(1, cell.Row - 250);
                for (int row = cell.Row - 1;
                    row >= firstRow;
                    row--)
                {
                    string rowType =
                        ReadText(sheet, row, 9);
                    string rowRateId =
                        ReadText(sheet, row, 10);
                    if (!string.Equals(
                        rowType,
                        "SECTION_TOTAL",
                        StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(
                            rowRateId,
                            rateId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    Excel.Range candidate = null;
                    try
                    {
                        candidate =
                            sheet.Cells[row, amountColumn]
                                as Excel.Range;
                        string formula =
                            Convert.ToString(
                                candidate?.Formula,
                                CultureInfo.InvariantCulture) ??
                            string.Empty;
                        if (!formula.StartsWith(
                            "=",
                            StringComparison.Ordinal))
                        {
                            continue;
                        }

                        string letter =
                            ExcelColumnAddress.ToLetters(
                                amountColumn);
                        expected =
                            "=SUM(" +
                            letter +
                            row.ToString(
                                CultureInfo.InvariantCulture) +
                            ":" +
                            letter +
                            row.ToString(
                                CultureInfo.InvariantCulture) +
                            ")";
                        return true;
                    }
                    finally
                    {
                        Release(candidate);
                    }
                }

                return false;
            }
            finally
            {
                Release(sheet);
            }
        }

        private static void ScanThkpFormulas(
            Excel.Workbook workbook,
            IDictionary<string, FormulaCell> overwrittenCells,
            IDictionary<string, ErrorCell> errorCells,
            ICollection<EstimateV2ValidationFinding> findings)
        {
            Excel.Worksheet sheet =
                FindWorksheet(
                    workbook,
                    "THKP-TC");
            if (sheet == null)
                return;

            try
            {
                int symbolColumn;
                int amountColumn;
                int materialRow;
                int laborRow;
                int machineRow;
                int totalRow;
                if (!TryFindThkpDirectLayout(
                    sheet,
                    out symbolColumn,
                    out amountColumn,
                    out materialRow,
                    out laborRow,
                    out machineRow,
                    out totalRow))
                {
                    findings.Add(new EstimateV2ValidationFinding(
                        "THKP_LAYOUT",
                        "Không nhận diện được layout THKP-TC",
                        "Không tìm đủ cột Ký hiệu/Thành tiền và các dòng VL/NC/M/T.",
                        EstimateV2CostIssueSeverity.Warning,
                        worksheetName: sheet.Name));
                    return;
                }

                int[] rows =
                {
                    materialRow,
                    laborRow,
                    machineRow,
                    totalRow
                };
                string[] expected =
                    EstimateV2ValidationRules
                        .ExpectedThkpDirectFormulas();

                for (int index = 0;
                    index < rows.Length;
                    index++)
                {
                    Excel.Range cell = null;
                    try
                    {
                        cell = sheet.Cells[
                            rows[index],
                            amountColumn] as Excel.Range;
                        string formula =
                            Convert.ToString(
                                cell.Formula,
                                CultureInfo.InvariantCulture) ??
                            string.Empty;
                        string address =
                            cell.Address[
                                false,
                                false,
                                Excel.XlReferenceStyle.xlA1,
                                false,
                                Type.Missing];
                        string key =
                            CellKey(
                                sheet,
                                rows[index],
                                amountColumn);

                        if (!EstimateV2ValidationRules
                            .FormulaEquivalent(
                                formula,
                                expected[index]))
                        {
                            overwrittenCells[key] =
                                new FormulaCell(
                                    sheet.Name,
                                    address,
                                    string.Empty,
                                    formula,
                                    expected[index],
                                    "THKP_FORMULA_OVERWRITTEN");
                        }
                        AddErrorIfAny(
                            errorCells,
                            sheet,
                            cell,
                            string.Empty);
                    }
                    finally
                    {
                        Release(cell);
                    }
                }
            }
            finally
            {
                Release(sheet);
            }
        }

        private static void ScanManagedSheetErrors(
            Excel.Workbook workbook,
            IDictionary<string, ErrorCell> errorCells)
        {
            foreach (string name in ManagedSheetNames)
            {
                Excel.Worksheet sheet =
                    FindWorksheet(
                        workbook,
                        name);
                if (sheet == null)
                    continue;

                Excel.Range used = null;
                try
                {
                    used = sheet.UsedRange;
                    int rows = used.Rows.Count;
                    int columns = used.Columns.Count;
                    if ((long)rows * columns > 100000L)
                        continue;

                    object values = used.Value2;
                    Array matrix = values as Array;
                    if (matrix == null)
                    {
                        AddErrorValue(
                            errorCells,
                            sheet,
                            used.Row,
                            used.Column,
                            values,
                            string.Empty);
                        continue;
                    }

                    int rowBase =
                        matrix.GetLowerBound(0);
                    int columnBase =
                        matrix.GetLowerBound(1);
                    for (int row = 0;
                        row < rows;
                        row++)
                    {
                        for (int column = 0;
                            column < columns;
                            column++)
                        {
                            object value =
                                matrix.GetValue(
                                    rowBase + row,
                                    columnBase + column);
                            AddErrorValue(
                                errorCells,
                                sheet,
                                used.Row + row,
                                used.Column + column,
                                value,
                                string.Empty);
                        }
                    }
                }
                finally
                {
                    Release(used);
                    Release(sheet);
                }
            }
        }

        private static void AddFormulaFindings(
            IReadOnlyDictionary<string, ErrorCell> errorCells,
            IReadOnlyDictionary<string, FormulaCell> overwrittenCells,
            ISet<string> warningWorkItems,
            ISet<string> formulaIssueWorkItems,
            ISet<string> overwrittenWorkItems,
            ICollection<EstimateV2ValidationFinding> findings)
        {
            foreach (IGrouping<string, ErrorCell> group in
                errorCells.Values.GroupBy(
                    item => item.ErrorText,
                    StringComparer.OrdinalIgnoreCase))
            {
                ErrorCell first = group.First();
                string[] affected = group
                    .Select(item => item.WorkItemId)
                    .Where(EstimateV2WorkItemState.IsValidId)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                foreach (string id in affected)
                {
                    formulaIssueWorkItems.Add(id);
                    warningWorkItems.Add(id);
                }

                EstimateV2CostIssueSeverity severity =
                    group.Key == "#REF!" ||
                    group.Key == "#VALUE!" ||
                    group.Key == "#N/A"
                        ? EstimateV2CostIssueSeverity.Error
                        : EstimateV2CostIssueSeverity.Warning;

                findings.Add(new EstimateV2ValidationFinding(
                    "MANAGED_FORMULA_ERROR",
                    "Phát hiện " +
                        group.Count().ToString("N0") +
                        " ô lỗi " + group.Key,
                    "Lỗi trong vùng dữ liệu/công thức do V2 quản lý.",
                    severity,
                    affected,
                    first.WorksheetName,
                    first.Address,
                    true));
            }

            foreach (IGrouping<string, FormulaCell> group in
                overwrittenCells.Values.GroupBy(
                    item => item.Code,
                    StringComparer.OrdinalIgnoreCase))
            {
                FormulaCell first = group.First();
                string[] affected = group
                    .Select(item => item.WorkItemId)
                    .Where(EstimateV2WorkItemState.IsValidId)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();

                foreach (string id in affected)
                {
                    overwrittenWorkItems.Add(id);
                    formulaIssueWorkItems.Add(id);
                    warningWorkItems.Add(id);
                }

                findings.Add(new EstimateV2ValidationFinding(
                    group.Key,
                    "Phát hiện " +
                        group.Count().ToString("N0") +
                        " công thức V2 bị ghi đè",
                    "Ô đầu tiên: " +
                        first.WorksheetName +
                        "!" +
                        first.Address +
                        ". Có thể khôi phục bằng chức năng cập nhật tương ứng.",
                    EstimateV2CostIssueSeverity.Error,
                    affected,
                    first.WorksheetName,
                    first.Address,
                    true));
            }
        }

        private static void AddErrorIfAny(
            IDictionary<string, ErrorCell> errors,
            Excel.Worksheet sheet,
            Excel.Range cell,
            string workItemId)
        {
            if (cell == null || sheet == null)
                return;
            AddErrorValue(
                errors,
                sheet,
                cell.Row,
                cell.Column,
                cell.Value2,
                workItemId);
        }

        private static void AddErrorValue(
            IDictionary<string, ErrorCell> errors,
            Excel.Worksheet sheet,
            int row,
            int column,
            object value,
            string workItemId)
        {
            int code;
            if (!TryGetExcelErrorCode(
                value,
                out code))
            {
                return;
            }

            string key =
                CellKey(
                    sheet,
                    row,
                    column);
            if (errors.ContainsKey(key))
            {
                if (EstimateV2WorkItemState.IsValidId(
                    workItemId) &&
                    !EstimateV2WorkItemState.IsValidId(
                        errors[key].WorkItemId))
                {
                    errors[key] =
                        errors[key].WithWorkItemId(
                            workItemId);
                }
                return;
            }

            errors[key] =
                new ErrorCell(
                    sheet.Name,
                    ExcelColumnAddress.ToLetters(column) +
                        row.ToString(
                            CultureInfo.InvariantCulture),
                    workItemId,
                    EstimateV2ValidationRules
                        .ExcelErrorText(code));
        }

        private static bool TryGetExcelErrorCode(
            object value,
            out int code)
        {
            ErrorWrapper wrapper =
                value as ErrorWrapper;
            if (wrapper != null)
            {
                code = wrapper.ErrorCode;
                return true;
            }

            if (value is int)
            {
                int raw = (int)value;
                uint encoded =
                    unchecked((uint)raw);
                if ((encoded & 0xFFFF0000u) ==
                    0x800A0000u)
                {
                    code = raw;
                    return true;
                }
            }

            code = 0;
            return false;
        }

        private static bool TryResolveNameRange(
            Excel.Workbook workbook,
            string name,
            out Excel.Range range)
        {
            Excel.Names names = null;
            Excel.Name defined = null;
            try
            {
                names = workbook.Names;
                try
                {
                    defined = names.Item(
                        name,
                        Type.Missing,
                        Type.Missing);
                }
                catch (COMException)
                {
                    range = null;
                    return false;
                }

                try
                {
                    range =
                        defined.RefersToRange;
                    return range != null;
                }
                catch (COMException)
                {
                    range = null;
                    return false;
                }
            }
            finally
            {
                Release(defined);
                Release(names);
            }
        }

        private static bool TryFindThkpDirectLayout(
            Excel.Worksheet sheet,
            out int symbolColumn,
            out int amountColumn,
            out int materialRow,
            out int laborRow,
            out int machineRow,
            out int totalRow)
        {
            symbolColumn = 0;
            amountColumn = 0;
            materialRow = 0;
            laborRow = 0;
            machineRow = 0;
            totalRow = 0;

            Excel.Range used = null;
            try
            {
                used = sheet.UsedRange;
                int lastRow = Math.Min(
                    80,
                    used.Row + used.Rows.Count - 1);
                int lastColumn = Math.Min(
                    14,
                    used.Column + used.Columns.Count - 1);
                int headerRow = 0;

                for (int row = 1;
                    row <= Math.Min(20, lastRow);
                    row++)
                {
                    int candidateSymbol = 0;
                    int candidateAmount = 0;
                    for (int column = 1;
                        column <= lastColumn;
                        column++)
                    {
                        string text = NormalizeHeader(
                            ReadText(
                                sheet,
                                row,
                                column));
                        if (text == "KY HIEU")
                            candidateSymbol = column;
                        if (text.StartsWith(
                            "THANH TIEN",
                            StringComparison.Ordinal))
                        {
                            candidateAmount = column;
                        }
                    }

                    if (candidateSymbol > 0 &&
                        candidateAmount > 0)
                    {
                        headerRow = row;
                        symbolColumn =
                            candidateSymbol;
                        amountColumn =
                            candidateAmount;
                        break;
                    }
                }

                if (headerRow == 0)
                    return false;

                for (int row = headerRow + 1;
                    row <= Math.Min(
                        lastRow,
                        headerRow + 40);
                    row++)
                {
                    string symbol =
                        NormalizeHeader(
                            ReadText(
                                sheet,
                                row,
                                symbolColumn));
                    if (symbol == "VL" &&
                        materialRow == 0)
                    {
                        materialRow = row;
                    }
                    else if (symbol == "NC" &&
                        laborRow == 0)
                    {
                        laborRow = row;
                    }
                    else if (symbol == "M" &&
                        machineRow == 0)
                    {
                        machineRow = row;
                    }
                    else if (symbol == "T" &&
                        totalRow == 0)
                    {
                        totalRow = row;
                    }
                }

                return materialRow > 0 &&
                    laborRow > 0 &&
                    machineRow > 0 &&
                    totalRow > 0;
            }
            finally
            {
                Release(used);
            }
        }

        private static string NormalizeHeader(
            string value)
        {
            string decomposed =
                (value ?? string.Empty)
                    .Normalize(
                        System.Text.NormalizationForm.FormD);
            var builder =
                new System.Text.StringBuilder();
            foreach (char ch in decomposed)
            {
                UnicodeCategory category =
                    CharUnicodeInfo.GetUnicodeCategory(ch);
                if (category !=
                    UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(
                        char.ToUpperInvariant(ch));
                }
            }

            return string.Join(
                " ",
                builder.ToString()
                    .Replace('Đ', 'D')
                    .Replace('đ', 'D')
                    .Split(
                        new[]
                        {
                            ' ', '\t', '\r', '\n',
                            '-', '_', '/', '(', ')'
                        },
                        StringSplitOptions
                            .RemoveEmptyEntries));
        }

        private static bool IsWorkItemRow(
            string workCode,
            string description,
            string unit,
            string quantity)
        {
            if (workCode.Length == 0 &&
                description.Length == 0)
            {
                return false;
            }
            return workCode.Length > 0 ||
                unit.Length > 0 ||
                quantity.Length > 0;
        }

        private static Excel.Worksheet FindWorksheet(
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
                        sheet =
                            sheets.Item[index]
                                as Excel.Worksheet;
                        if (sheet != null &&
                            string.Equals(
                                sheet.Name,
                                name,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Worksheet result =
                                sheet;
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

        private static string ReadText(
            Excel.Worksheet sheet,
            int row,
            int column)
        {
            Excel.Range cell = null;
            try
            {
                cell = sheet.Cells[row, column]
                    as Excel.Range;
                return (Convert.ToString(
                    cell?.Value2,
                    CultureInfo.InvariantCulture) ??
                    string.Empty).Trim();
            }
            finally
            {
                Release(cell);
            }
        }

        private static string CellKey(
            Excel.Worksheet sheet,
            int row,
            int column)
        {
            string codeName =
                (sheet.CodeName ?? string.Empty).Trim();
            if (codeName.Length == 0)
                codeName =
                    (sheet.Name ?? string.Empty).Trim();
            return codeName + "|" +
                row.ToString(
                    CultureInfo.InvariantCulture) +
                "|" +
                column.ToString(
                    CultureInfo.InvariantCulture);
        }

        private static void Release(
            object value)
        {
            if (value != null &&
                Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }

        private sealed class ErrorCell
        {
            internal ErrorCell(
                string worksheetName,
                string address,
                string workItemId,
                string errorText)
            {
                WorksheetName =
                    worksheetName ?? string.Empty;
                Address =
                    address ?? string.Empty;
                WorkItemId =
                    workItemId ?? string.Empty;
                ErrorText =
                    errorText ?? string.Empty;
            }

            internal string WorksheetName { get; }
            internal string Address { get; }
            internal string WorkItemId { get; }
            internal string ErrorText { get; }

            internal ErrorCell WithWorkItemId(
                string workItemId)
            {
                return new ErrorCell(
                    WorksheetName,
                    Address,
                    workItemId,
                    ErrorText);
            }
        }

        private sealed class FormulaCell
        {
            internal FormulaCell(
                string worksheetName,
                string address,
                string workItemId,
                string actual,
                string expected,
                string code)
            {
                WorksheetName =
                    worksheetName ?? string.Empty;
                Address = address ?? string.Empty;
                WorkItemId =
                    workItemId ?? string.Empty;
                Actual = actual ?? string.Empty;
                Expected = expected ?? string.Empty;
                Code = code ?? string.Empty;
            }

            internal string WorksheetName { get; }
            internal string Address { get; }
            internal string WorkItemId { get; }
            internal string Actual { get; }
            internal string Expected { get; }
            internal string Code { get; }
        }
    }
}
