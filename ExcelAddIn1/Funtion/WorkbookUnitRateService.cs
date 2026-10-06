using ExcelAddIn1.Core;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookUnitRateContext
    {
        internal WorkbookUnitRateContext(
            string packageId,
            string packageVersion,
            string packageChecksum,
            NormCatalog normCatalog,
            PriceProfile priceProfile)
        {
            PackageId = packageId;
            PackageVersion = packageVersion;
            PackageChecksum = packageChecksum;
            NormCatalog = normCatalog;
            PriceProfile = priceProfile;
        }

        public string PackageId { get; }
        public string PackageVersion { get; }
        public string PackageChecksum { get; }
        public NormCatalog NormCatalog { get; }
        public PriceProfile PriceProfile { get; }
    }

    public static class WorkbookUnitRateService
    {
        public static WorkbookUnitRateContext LoadRequired(Excel.Workbook workbook)
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
            PriceProfile priceProfile = WorkbookPriceProfileService.LoadRequired(workbook);
            return new WorkbookUnitRateContext(
                bundle.Package.PackageId,
                bundle.Package.DataVersion,
                bundle.Package.PackageChecksum,
                catalog,
                priceProfile);
        }

        public static UnitRateCalculationResult Calculate(
            WorkbookUnitRateContext context,
            string normKey,
            string variantCode,
            string[] conditions,
            UnitRateResourceBinding[] bindings)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            return UnitRateCalculator.Calculate(new UnitRateCalculationRequest(
                context.NormCatalog.FindRequired(normKey),
                variantCode,
                context.PriceProfile,
                conditions,
                bindings));
        }
    }
}
