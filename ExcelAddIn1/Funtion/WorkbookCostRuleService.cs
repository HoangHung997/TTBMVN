using ExcelAddIn1.Core;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookCostRuleContext
    {
        internal WorkbookCostRuleContext(RegulationPackage package, CostRuleCatalog catalog)
        {
            PackageId = package.PackageId;
            PackageVersion = package.DataVersion;
            PackageChecksum = package.PackageChecksum;
            Catalog = catalog;
        }

        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public CostRuleCatalog Catalog { get; }
    }

    public static class WorkbookCostRuleService
    {
        public static WorkbookCostRuleContext LoadPinnedCatalog(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            RegulationPackageBootstrapService.LoadAvailablePackages();
            ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            RegulationPackageBundle bundle = store.LoadBundleRequired(
                profile.RegulationPackageId,
                profile.RegulationPackageVersion,
                profile.RegulationPackageChecksum);
            RegulationDataModule module;
            if (!bundle.Modules.TryGetValue(RegulationModuleKind.CostRule, out module))
                throw new InvalidOperationException("Package đã pin không có module CostRule.");

            try
            {
                return new WorkbookCostRuleContext(bundle.Package, CostRuleCatalog.Load(module));
            }
            catch (Exception ex) when (
                ex is ArgumentException ||
                ex is FormatException ||
                ex is InvalidOperationException)
            {
                throw new InvalidOperationException(
                    "Không thể nạp bảng chi phí từ package " +
                    bundle.Package.PackageId + " v" + bundle.Package.DataVersion + ". " + ex.Message,
                    ex);
            }
        }
    }
}
