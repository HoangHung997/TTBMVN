using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookEstimateCalculationContext
    {
        internal WorkbookEstimateCalculationContext(
            string packageId,
            string packageVersion,
            string packageChecksum,
            NormCatalog normCatalog,
            PriceProfilePortfolio pricePortfolio)
        {
            PackageId = packageId;
            PackageVersion = packageVersion;
            PackageChecksum = packageChecksum;
            NormCatalog = normCatalog;
            PricePortfolio = pricePortfolio;
        }

        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public NormCatalog NormCatalog { get; }
        public PriceProfilePortfolio PricePortfolio { get; }
    }

    public sealed class WorkbookEstimateRatePreview
    {
        internal WorkbookEstimateRatePreview(
            EstimateRateGroup group,
            UnitRateCalculationResult result,
            string error)
        {
            Group = group;
            Result = result;
            Error = error ?? string.Empty;
        }

        public EstimateRateGroup Group { get; }
        public UnitRateCalculationResult Result { get; }
        public string Error { get; }
        public bool IsValid => Result != null && Error.Length == 0;
    }

    public sealed class WorkbookEstimateWorkspacePreview
    {
        internal WorkbookEstimateWorkspacePreview(
            EstimateWorkspace workspace,
            WorkbookEstimateCalculationContext context,
            EstimateRatePlan plan,
            IEnumerable<WorkbookEstimateRatePreview> rates)
        {
            Workspace = workspace;
            Context = context;
            Plan = plan;
            Rates = new ReadOnlyCollection<WorkbookEstimateRatePreview>(rates.ToList());
        }

        public EstimateWorkspace Workspace { get; }
        public WorkbookEstimateCalculationContext Context { get; }
        public EstimateRatePlan Plan { get; }
        public IReadOnlyList<WorkbookEstimateRatePreview> Rates { get; }
        public bool IsValid => Plan.IsValid && Rates.All(rate => rate.IsValid);
        public int CalculatedRowCount => Rates.Sum(rate => rate.Group.Rows.Count);

        public WorkbookEstimateRatePreview FindRateForRow(string rowId)
        {
            return Rates.SingleOrDefault(rate => rate.Group.Rows.Any(row => string.Equals(
                row.RowId,
                rowId,
                StringComparison.OrdinalIgnoreCase)));
        }
    }

    public static class WorkbookEstimateCalculationService
    {
        public static WorkbookEstimateCalculationContext LoadContext(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            RegulationPackageBootstrapService.LoadAvailablePackages();
            ProjectProfile project = WorkbookProjectProfileService.LoadRequired(workbook);
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            RegulationPackageBundle bundle = store.LoadBundleRequired(
                project.RegulationPackageId,
                project.RegulationPackageVersion,
                project.RegulationPackageChecksum);
            RegulationDataModule module;
            if (!bundle.Modules.TryGetValue(RegulationModuleKind.Norm, out module))
                throw new InvalidOperationException("Package da pin khong co module dinh muc.");
            NormCatalog catalog;
            try
            {
                catalog = NormCatalog.Load(module);
            }
            catch (FormatException ex)
            {
                throw new InvalidOperationException(
                    "Package da pin chua co bang hao phi chi tiet de tinh don gia.",
                    ex);
            }
            PriceProfilePortfolio portfolio;
            WorkbookPriceProfilePortfolioService.TryLoad(workbook, out portfolio);
            return new WorkbookEstimateCalculationContext(
                bundle.Package.PackageId,
                bundle.Package.DataVersion,
                bundle.Package.PackageChecksum,
                catalog,
                portfolio);
        }

        public static WorkbookEstimateWorkspacePreview Preview(
            Excel.Workbook workbook,
            EstimateWorkspace workspace)
        {
            WorkbookEstimateCalculationContext context = LoadContext(workbook);
            PriceProfile[] profiles = context.PricePortfolio?.Profiles.ToArray() ?? Array.Empty<PriceProfile>();
            EstimateRatePlan plan = EstimateRatePlan.Build(
                workspace,
                context.PackageId,
                context.PackageVersion,
                context.PackageChecksum,
                profiles);
            var rates = new List<WorkbookEstimateRatePreview>();
            foreach (EstimateRateGroup group in plan.Groups)
            {
                try
                {
                    PriceProfile profile = context.PricePortfolio.FindRequired(group.Identity.LaborAudience);
                    NormDefinition definition = context.NormCatalog.FindRequired(group.Identity.NormKey);
                    UnitRateCalculationResult result = UnitRateCalculator.Calculate(
                        new UnitRateCalculationRequest(
                            definition,
                            group.Identity.VariantCode,
                            profile,
                            group.Identity.Conditions,
                            group.Identity.Bindings));
                    rates.Add(new WorkbookEstimateRatePreview(group, result, string.Empty));
                }
                catch (Exception ex) when (
                    ex is ArgumentException ||
                    ex is InvalidOperationException ||
                    ex is KeyNotFoundException)
                {
                    rates.Add(new WorkbookEstimateRatePreview(group, null, ex.Message));
                }
            }
            return new WorkbookEstimateWorkspacePreview(workspace, context, plan, rates);
        }
    }
}
