using ExcelAddIn1.Core;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookRegulationPackageService
    {
        public static bool PinAndSave(
            Excel.Workbook workbook,
            RegulationPackage package)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            ProjectProfile current = WorkbookProjectProfileService.LoadRequired(workbook);
            ProjectProfile pinned = RegulationPackagePinService.Pin(current, package);
            return WorkbookProjectProfileService.Save(workbook, pinned);
        }

        public static RegulationPackagePinVerificationResult Verify(
            Excel.Workbook workbook,
            string storeRoot)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
            var store = new RegulationPackageStore(storeRoot);
            return RegulationPackagePinService.Verify(profile, store);
        }

        public static RegulationPackagePinVerificationResult VerifyDefaultStore(
            Excel.Workbook workbook)
        {
            return Verify(workbook, AppPaths.RegulationPackageDirectory);
        }

        public static RegulationPackage LoadPinnedPackageRequired(
            Excel.Workbook workbook,
            string storeRoot)
        {
            RegulationPackagePinVerificationResult result = Verify(workbook, storeRoot);
            if (!result.IsAvailable)
                throw new InvalidOperationException(result.Message);
            return result.Package;
        }
    }
}
