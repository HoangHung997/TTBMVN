using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public enum WorkbookPackageMigrationStatus
    {
        Canceled,
        NoChanges,
        Applied
    }

    public enum WorkbookPackageMigrationPhase
    {
        BackupCreated,
        ProfilePinned,
        EstimateWritten,
        CostSummaryWritten,
        ValidationCompleted,
        WorkbookSaved
    }

    public sealed class WorkbookPackageMigrationException : InvalidOperationException
    {
        public WorkbookPackageMigrationException(
            string message,
            string backupPath,
            bool rollbackAttempted,
            bool rollbackSucceeded,
            Exception innerException)
            : base(message, innerException)
        {
            BackupPath = backupPath ?? string.Empty;
            RollbackAttempted = rollbackAttempted;
            RollbackSucceeded = rollbackSucceeded;
        }

        public string BackupPath { get; }
        public bool RollbackAttempted { get; }
        public bool RollbackSucceeded { get; }
    }

    public static class WorkbookPackageMigrationService
    {
        public static WorkbookPackageMigrationPreview Preview(
            Excel.Workbook workbook,
            string storeRoot,
            string targetPackageId,
            string targetDataVersion,
            string targetChecksum)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            var store = new RegulationPackageStore(storeRoot);
            ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
            RegulationPackagePinVerificationResult current =
                RegulationPackagePinService.Verify(profile, store);
            if (!current.IsAvailable)
                throw new InvalidOperationException(current.Message);
            RegulationPackage target = store.LoadRequired(
                targetPackageId,
                targetDataVersion,
                targetChecksum);
            RegulationPackageMigrationPlan plan =
                RegulationPackageMigrationPlan.Create(current.Package, target);
            if (!plan.Diff.HasChanges)
            {
                return new WorkbookPackageMigrationPreview(
                    profile.Clone(),
                    plan,
                    new WorkbookPackageMigrationImpact(
                        string.Empty,
                        string.Empty,
                        string.Empty,
                        0,
                        0,
                        new string[0],
                        new string[0],
                        new WorkbookPackageMigrationValueChange[0],
                        string.Empty),
                    null,
                    null,
                    string.Empty);
            }

            return BuildCalculationPreview(workbook, store, profile, plan);
        }

        public static WorkbookPackageMigrationResult Apply(
            Excel.Workbook workbook,
            string storeRoot,
            WorkbookPackageMigrationPreview preview,
            string backupPath,
            bool confirmed)
        {
            return Apply(workbook, storeRoot, preview, backupPath, confirmed, null);
        }

        public static WorkbookPackageMigrationResult Apply(
            Excel.Workbook workbook,
            string storeRoot,
            WorkbookPackageMigrationPreview preview,
            string backupPath,
            bool confirmed,
            Action<WorkbookPackageMigrationPhase> phaseCallback)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (preview == null)
                throw new ArgumentNullException(nameof(preview));
            if (!confirmed)
            {
                return new WorkbookPackageMigrationResult(
                    WorkbookPackageMigrationStatus.Canceled,
                    preview.Plan,
                    preview.Impact,
                    string.Empty,
                    null);
            }
            if (!preview.Plan.Diff.HasChanges)
            {
                return new WorkbookPackageMigrationResult(
                    WorkbookPackageMigrationStatus.NoChanges,
                    preview.Plan,
                    preview.Impact,
                    string.Empty,
                    null);
            }
            if (preview.Impact == null || !preview.Impact.CanApply ||
                preview.TargetEstimate == null || preview.TargetCostSummary == null)
            {
                throw new InvalidOperationException(
                    "Preview migration con loi; khong duoc thay doi workbook. " +
                    string.Join(" ", preview.Impact?.Issues ?? new string[0]));
            }

            string normalizedBackupPath = ValidateBackupPath(workbook, backupPath);
            var store = new RegulationPackageStore(storeRoot);
            WorkbookPackageMigrationSnapshot snapshot = null;
            bool mutationStarted = false;
            bool rollbackAttempted = false;
            bool rollbackSucceeded = false;
            try
            {
                workbook.Save();
                WorkbookPackageMigrationPreview fresh = Preview(
                    workbook,
                    storeRoot,
                    preview.Plan.Target.PackageId,
                    preview.Plan.Target.DataVersion,
                    preview.Plan.Target.PackageChecksum);
                EnsurePreviewStillCurrent(preview, fresh);
                snapshot = WorkbookPackageMigrationSnapshot.Capture(
                    workbook,
                    fresh.TargetEstimate.Plan);
                if (!string.Equals(
                    snapshot.Fingerprint,
                    fresh.WorkbookFingerprint,
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Workbook da doi trong luc tao snapshot migration.");
                }

                Directory.CreateDirectory(Path.GetDirectoryName(normalizedBackupPath));
                workbook.SaveCopyAs(normalizedBackupPath);
                phaseCallback?.Invoke(WorkbookPackageMigrationPhase.BackupCreated);

                mutationStarted = true;
                WorkbookValidationReport validation;
                using (new ExcelWriteContext(workbook.Application))
                {
                    ProjectProfile currentProfile = WorkbookProjectProfileService.LoadRequired(workbook);
                    EnsureProfileMatchesPackage(
                        currentProfile,
                        fresh.Plan.Source,
                        "Workbook da doi package sau luc preview.");
                    RegulationPackage target = store.LoadRequired(
                        fresh.Plan.Target.PackageId,
                        fresh.Plan.Target.DataVersion,
                        fresh.Plan.Target.PackageChecksum);
                    ProjectProfile targetProfile = RegulationPackagePinService.Pin(currentProfile, target);
                    WorkbookProjectProfileService.Save(workbook, targetProfile);
                    phaseCallback?.Invoke(WorkbookPackageMigrationPhase.ProfilePinned);

                    WorkbookEstimateAppendixWriter.Apply(workbook, fresh.TargetEstimate);
                    phaseCallback?.Invoke(WorkbookPackageMigrationPhase.EstimateWritten);

                    WorkbookCostSummaryWriter.Apply(workbook, fresh.TargetCostSummary);
                    phaseCallback?.Invoke(WorkbookPackageMigrationPhase.CostSummaryWritten);

                    validation = WorkbookValidationService.Scan(workbook);
                    if (!validation.IsValid)
                    {
                        string errors = string.Join(" | ", validation.Issues
                            .Where(issue => issue.IsBlocking)
                            .Take(5)
                            .Select(issue => issue.Code + ": " + issue.Message));
                        throw new InvalidOperationException(
                            "Validation sau migration khong dat. " + errors);
                    }
                    phaseCallback?.Invoke(WorkbookPackageMigrationPhase.ValidationCompleted);
                }

                // Excel persists Calculation in the workbook, so save after the context
                // has restored the user's application state.
                workbook.Save();
                phaseCallback?.Invoke(WorkbookPackageMigrationPhase.WorkbookSaved);

                RegulationPackagePinVerificationResult verified =
                    RegulationPackagePinService.Verify(
                        WorkbookProjectProfileService.LoadRequired(workbook),
                        store);
                if (!verified.IsAvailable ||
                    !string.Equals(
                        verified.Package.PackageChecksum,
                        fresh.Plan.Target.PackageChecksum,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Verify package target sau migration khong dat.");
                }

                return new WorkbookPackageMigrationResult(
                    WorkbookPackageMigrationStatus.Applied,
                    fresh.Plan,
                    fresh.Impact,
                    normalizedBackupPath,
                    validation);
            }
            catch (Exception migrationException)
            {
                Exception finalException = migrationException;
                if (mutationStarted && snapshot != null)
                {
                    rollbackAttempted = true;
                    try
                    {
                        using (new ExcelWriteContext(workbook.Application))
                        {
                            snapshot.Restore(workbook);
                            workbook.Application.Calculate();
                            snapshot.VerifyRestored(
                                workbook,
                                preview.TargetEstimate.Plan);
                        }
                        workbook.Save();
                        ProjectProfile restored = WorkbookProjectProfileService.LoadRequired(workbook);
                        EnsureProfileMatchesPackage(
                            restored,
                            preview.Plan.Source,
                            "Rollback package pin khong khop source.");
                        snapshot.VerifyRestored(workbook, preview.TargetEstimate.Plan);
                        rollbackSucceeded = true;
                    }
                    catch (Exception rollbackException)
                    {
                        RuntimeLogger.LogOperation(
                            rollbackException,
                            "Rollback workbook package migration",
                            "DT-504",
                            "rollback",
                            workbook,
                            string.Empty);
                        finalException = new AggregateException(migrationException, rollbackException);
                    }
                }

                RuntimeLogger.LogOperation(
                    finalException,
                    "Apply workbook package migration",
                    "DT-504",
                    "apply",
                    workbook,
                    string.Empty);
                throw new WorkbookPackageMigrationException(
                    rollbackAttempted && rollbackSucceeded
                        ? "Migration loi; profile, bang gia, audit va ket qua cu da duoc khoi phuc."
                        : "Migration package that bai; su dung file backup neu workbook khong con nhat quan.",
                    normalizedBackupPath,
                    rollbackAttempted,
                    rollbackSucceeded,
                    finalException);
            }
        }

        private static WorkbookPackageMigrationPreview BuildCalculationPreview(
            Excel.Workbook workbook,
            RegulationPackageStore store,
            ProjectProfile profile,
            RegulationPackageMigrationPlan plan)
        {
            PriceProfile priceProfile = WorkbookPriceProfileService.LoadRequired(workbook);
            RegulationPackageBundle targetBundle = store.LoadBundleRequired(
                plan.Target.PackageId,
                plan.Target.DataVersion,
                plan.Target.PackageChecksum);
            WorkbookUnitRateContext targetUnitRate = BuildUnitRateContext(targetBundle, priceProfile);

            Excel.Worksheet estimateSheet = null;
            Excel.Worksheet summarySheet = null;
            try
            {
                estimateSheet = WorksheetRoleService.ResolveRequired(
                    workbook,
                    WorksheetRole.EstimateAppendix);
                summarySheet = WorksheetRoleService.ResolveRequired(
                    workbook,
                    WorksheetRole.CostSummary);
                WorkbookEstimatePlan estimatePlan;
                if (!WorkbookResultAuditService.TryRestoreEstimatePlanForValidation(
                    workbook,
                    estimateSheet,
                    targetUnitRate,
                    out estimatePlan))
                {
                    estimatePlan = WorkbookEstimateAppendixService.ImportLegacyPlan(
                        workbook,
                        targetUnitRate);
                }
                IReadOnlyList<string> conditions = WorkbookResultAuditService.ReadEstimateConditions(
                    workbook,
                    estimateSheet.CodeName);
                WorkbookEstimatePreview targetEstimate = WorkbookEstimateAppendixService.Preview(
                    targetUnitRate,
                    estimatePlan,
                    conditions);
                var issues = new List<string>();
                AddEstimateIssues(issues, "Package dich", targetEstimate);

                WorkbookCostSummaryPreview targetCost = null;
                if (targetEstimate.IsValid)
                {
                    try
                    {
                        WorkbookCostSummaryContext currentCost = WorkbookCostSummaryService.Load(workbook);
                        CostSummaryCalculationRequest targetRequest = WithEstimateTotals(
                            currentCost.DefaultRequest,
                            targetEstimate.Result);
                        WorkbookCostSummaryContext targetCostContext = new WorkbookCostSummaryContext(
                            plan.Target.PackageId,
                            plan.Target.DataVersion,
                            plan.Target.PackageChecksum,
                            LoadCostCatalog(targetBundle),
                            currentCost.SummarySheetCodeName,
                            currentCost.EstimateSheetCodeName,
                            targetRequest,
                            currentCost.ImportNote);
                        targetCost = WorkbookCostSummaryService.Preview(targetCostContext, targetRequest);
                    }
                    catch (Exception ex) when (
                        ex is ArgumentException ||
                        ex is FormatException ||
                        ex is InvalidOperationException)
                    {
                        issues.Add("Khong tinh duoc tong hop kinh phi package dich: " + ex.Message);
                    }
                }

                ResultAuditTrail audit = WorkbookResultAuditService.Load(workbook);
                string[] scopes = audit.Entries
                    .Select(entry => entry.ScopeId)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToArray();
                if (scopes.Length == 0)
                {
                    scopes = new[]
                    {
                        WorkbookResultAuditService.EstimateScopeId,
                        WorkbookResultAuditService.CostSummaryScopeId
                    };
                }

                WorkbookPackageMigrationSnapshot state =
                    WorkbookPackageMigrationSnapshot.Capture(workbook, estimatePlan);
                decimal[] sourceEstimateTotals = ReadEstimateTotals(estimateSheet);
                decimal sourceCostTotal = ReadDecimal(summarySheet, "E27");
                var changes = BuildValueChanges(
                    sourceEstimateTotals,
                    targetEstimate,
                    sourceCostTotal,
                    targetCost);
                string impactFingerprint = BuildImpactFingerprint(
                    plan,
                    priceProfile,
                    audit,
                    estimatePlan,
                    conditions,
                    changes,
                    issues,
                    state.Fingerprint);
                var impact = new WorkbookPackageMigrationImpact(
                    priceProfile.ProfileId,
                    priceProfile.DataVersion,
                    priceProfile.Checksum,
                    audit.Entries.Count,
                    estimatePlan.Lines.Count,
                    scopes,
                    issues,
                    changes,
                    impactFingerprint);
                return new WorkbookPackageMigrationPreview(
                    profile.Clone(),
                    plan,
                    impact,
                    targetEstimate,
                    targetCost,
                    state.Fingerprint);
            }
            finally
            {
                Release(summarySheet);
                Release(estimateSheet);
            }
        }

        private static WorkbookUnitRateContext BuildUnitRateContext(
            RegulationPackageBundle bundle,
            PriceProfile priceProfile)
        {
            RegulationDataModule module;
            if (!bundle.Modules.TryGetValue(RegulationModuleKind.Norm, out module))
                throw new InvalidOperationException("Package " + bundle.Package.PackageId + " thieu module Norm.");
            return new WorkbookUnitRateContext(
                bundle.Package.PackageId,
                bundle.Package.DataVersion,
                bundle.Package.PackageChecksum,
                NormCatalog.Load(module),
                priceProfile);
        }

        private static CostRuleCatalog LoadCostCatalog(RegulationPackageBundle bundle)
        {
            RegulationDataModule module;
            if (!bundle.Modules.TryGetValue(RegulationModuleKind.CostRule, out module))
                throw new InvalidOperationException("Package " + bundle.Package.PackageId + " thieu module CostRule.");
            return CostRuleCatalog.Load(module);
        }

        private static CostSummaryCalculationRequest WithEstimateTotals(
            CostSummaryCalculationRequest request,
            EstimateAppendixCalculationResult estimate)
        {
            return new CostSummaryCalculationRequest(
                request.Template,
                estimate.MaterialAmountVnd,
                estimate.LaborAmountVnd,
                estimate.MachineAmountVnd,
                request.Terrain,
                request.AreaHa,
                request.ProjectKind,
                request.ConstructionKind,
                request.DisposalWeightKg,
                request.PreTaxIncomeRatePercent,
                request.VatRatePercent,
                request.SelectedComponents,
                request.Overrides);
        }

        private static void AddEstimateIssues(
            ICollection<string> issues,
            string prefix,
            WorkbookEstimatePreview preview)
        {
            foreach (WorkbookEstimatePreviewLine line in preview.Lines.Where(line => line.IsBlocking))
            {
                issues.Add(
                    prefix + " khong tinh duoc hang " +
                    line.Line.TargetRow.ToString(CultureInfo.InvariantCulture) + " (" +
                    line.Line.NormKey + "): " + line.Error);
            }
            if (!preview.IsValid && !preview.Lines.Any(line => line.IsBlocking))
                issues.Add(prefix + " khong tao duoc ket qua phu luc.");
        }

        private static IReadOnlyList<WorkbookPackageMigrationValueChange> BuildValueChanges(
            decimal[] sourceEstimate,
            WorkbookEstimatePreview targetEstimate,
            decimal sourceCost,
            WorkbookCostSummaryPreview targetCost)
        {
            var changes = new List<WorkbookPackageMigrationValueChange>();
            if (sourceEstimate != null && sourceEstimate.Length == 6 && targetEstimate.Result != null)
            {
                changes.Add(Change("Estimate.Material", "Phu luc - Vat lieu",
                    sourceEstimate[0],
                    targetEstimate.Result.MaterialAmountVnd));
                changes.Add(Change("Estimate.Labor", "Phu luc - Nhan cong",
                    sourceEstimate[1],
                    targetEstimate.Result.LaborAmountVnd));
                changes.Add(Change("Estimate.Machine", "Phu luc - May",
                    sourceEstimate[2],
                    targetEstimate.Result.MachineAmountVnd));
                changes.Add(Change("Estimate.AcceptedMaterial", "Nghiem thu - Vat lieu",
                    sourceEstimate[3],
                    targetEstimate.Result.AcceptedMaterialAmountVnd));
                changes.Add(Change("Estimate.AcceptedLabor", "Nghiem thu - Nhan cong",
                    sourceEstimate[4],
                    targetEstimate.Result.AcceptedLaborAmountVnd));
                changes.Add(Change("Estimate.AcceptedMachine", "Nghiem thu - May",
                    sourceEstimate[5],
                    targetEstimate.Result.AcceptedMachineAmountVnd));
            }
            if (targetCost != null)
            {
                changes.Add(Change("CostSummary.RoundedAfterTax", "Tong kinh phi lam tron",
                    sourceCost,
                    targetCost.Result.RoundedAfterTaxVnd));
            }
            return changes.AsReadOnly();
        }

        private static decimal[] ReadEstimateTotals(Excel.Worksheet worksheet)
        {
            Excel.Range range = null;
            try
            {
                range = worksheet.Range["I27:N27"];
                Array values = range.Value2 as Array;
                if (values == null)
                    throw new InvalidOperationException("Khong doc duoc tong phu luc I27:N27.");
                var result = new decimal[6];
                int columnLower = values.GetLowerBound(1);
                for (int index = 0; index < result.Length; index++)
                {
                    result[index] = Convert.ToDecimal(
                        values.GetValue(values.GetLowerBound(0), columnLower + index) ?? 0d,
                        CultureInfo.InvariantCulture);
                }
                return result;
            }
            finally
            {
                Release(range);
            }
        }

        private static decimal ReadDecimal(Excel.Worksheet worksheet, string address)
        {
            Excel.Range range = null;
            try
            {
                range = worksheet.Range[address];
                return Convert.ToDecimal(range.Value2 ?? 0d, CultureInfo.InvariantCulture);
            }
            finally
            {
                Release(range);
            }
        }

        private static WorkbookPackageMigrationValueChange Change(
            string key,
            string label,
            decimal source,
            decimal target)
        {
            return new WorkbookPackageMigrationValueChange(key, label, source, target);
        }

        private static string BuildImpactFingerprint(
            RegulationPackageMigrationPlan plan,
            PriceProfile profile,
            ResultAuditTrail audit,
            WorkbookEstimatePlan estimatePlan,
            IEnumerable<string> conditions,
            IEnumerable<WorkbookPackageMigrationValueChange> changes,
            IEnumerable<string> issues,
            string workbookFingerprint)
        {
            var text = new StringBuilder();
            text.AppendLine(plan.PlanId);
            text.AppendLine(plan.Source.PackageChecksum);
            text.AppendLine(plan.Target.PackageChecksum);
            text.AppendLine(profile.Checksum);
            text.AppendLine(ProjectProfileSerializer.ComputeChecksum(
                ResultAuditTrailSerializer.Serialize(audit)));
            text.AppendLine(workbookFingerprint);
            foreach (WorkbookEstimateLine line in estimatePlan.Lines)
            {
                text.Append(line.TargetRow).Append('|').Append(line.LineId).Append('|')
                    .Append(line.NormKey).Append('|').Append(line.VariantCode).Append('|')
                    .Append(line.Quantity.ToString(CultureInfo.InvariantCulture)).Append('|')
                    .Append(line.AcceptedQuantity.ToString(CultureInfo.InvariantCulture)).AppendLine();
            }
            foreach (string condition in conditions)
                text.Append("condition=").AppendLine(condition);
            foreach (WorkbookPackageMigrationValueChange change in changes)
            {
                text.Append(change.Key).Append('|')
                    .Append(change.SourceValue.ToString(CultureInfo.InvariantCulture)).Append('|')
                    .Append(change.TargetValue.ToString(CultureInfo.InvariantCulture)).AppendLine();
            }
            foreach (string issue in issues)
                text.Append("issue=").AppendLine(issue);
            return ProjectProfileSerializer.ComputeChecksum(text.ToString());
        }

        private static void EnsurePreviewStillCurrent(
            WorkbookPackageMigrationPreview expected,
            WorkbookPackageMigrationPreview actual)
        {
            if (!string.Equals(expected.Plan.PlanId, actual.Plan.PlanId, StringComparison.Ordinal) ||
                !string.Equals(
                    expected.Impact.Fingerprint,
                    actual.Impact.Fingerprint,
                    StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(
                    expected.WorkbookFingerprint,
                    actual.WorkbookFingerprint,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "Workbook, bang gia hoac package store da doi sau luc preview. Hay preview lai.");
            }
            if (!actual.Impact.CanApply)
                throw new InvalidOperationException(string.Join(" ", actual.Impact.Issues));
        }

        private static string ValidateBackupPath(Excel.Workbook workbook, string backupPath)
        {
            if (string.IsNullOrWhiteSpace(backupPath))
                throw new ArgumentException("Backup path la bat buoc khi apply migration.", nameof(backupPath));
            string fullPath = Path.GetFullPath(backupPath);
            if (string.Equals(fullPath, workbook.FullName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Backup path khong duoc trung working workbook.", nameof(backupPath));
            if (File.Exists(fullPath) || Directory.Exists(fullPath))
                throw new IOException("Backup path da ton tai: " + fullPath);
            string directory = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(directory))
                throw new ArgumentException("Backup path khong co thu muc hop le.", nameof(backupPath));
            return fullPath;
        }

        private static void EnsureProfileMatchesPackage(
            ProjectProfile profile,
            RegulationPackage package,
            string errorMessage)
        {
            if (!string.Equals(profile.RegulationPackageId, package.PackageId, StringComparison.Ordinal) ||
                !string.Equals(profile.RegulationPackageVersion, package.DataVersion, StringComparison.Ordinal) ||
                !string.Equals(
                    profile.RegulationPackageChecksum,
                    package.PackageChecksum,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(errorMessage);
            }
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
        }
    }
}
