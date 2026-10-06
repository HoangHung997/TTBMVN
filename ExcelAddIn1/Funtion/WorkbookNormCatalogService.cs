using ExcelAddIn1.Core;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookNormCatalogService
    {
        public static NormSearchIndex LoadPinnedSearchIndex(Excel.Workbook workbook)
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
            RegulationDataModule normModule;
            if (!bundle.Modules.TryGetValue(RegulationModuleKind.Norm, out normModule))
                throw new InvalidOperationException("Package da pin khong co module Norm.");
            return new NormSearchIndex(
                bundle.Package.PackageId,
                bundle.Package.DataVersion,
                bundle.Package.PackageChecksum,
                normModule);
        }
    }
}
