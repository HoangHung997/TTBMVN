using ExcelAddIn1.Core;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookPriceProfilePortfolioService
    {
        public const string ManifestPropertyName = "TTBMVN.PricePortfolio.Manifest";
        private const string PartPropertyPrefix = "TTBMVN.PricePortfolio.Part.";
        private const string Label = "PriceProfilePortfolio";

        public static bool TryLoad(Excel.Workbook workbook, out PriceProfilePortfolio portfolio)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            string payload;
            if (WorkbookCustomPayloadStore.TryRead(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                Label,
                out payload))
            {
                portfolio = PriceProfilePortfolioSerializer.Deserialize(payload);
                return true;
            }

            PriceProfile legacy;
            if (WorkbookPriceProfileService.TryLoad(workbook, out legacy))
            {
                portfolio = PriceProfilePortfolio.Create(new[] { legacy });
                return true;
            }
            portfolio = null;
            return false;
        }

        public static PriceProfilePortfolio LoadRequired(Excel.Workbook workbook)
        {
            PriceProfilePortfolio portfolio;
            if (!TryLoad(workbook, out portfolio))
                throw new InvalidOperationException("Workbook chua co ho so gia HLNS/KHLNS.");
            return portfolio;
        }

        public static bool Save(Excel.Workbook workbook, PriceProfilePortfolio portfolio)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            PriceProfilePortfolioValidator.ValidateRequired(portfolio);
            return WorkbookCustomPayloadStore.Save(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                PriceProfilePortfolioSerializer.Serialize(portfolio),
                Label);
        }

        public static bool SaveProfile(Excel.Workbook workbook, PriceProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            PriceProfilePortfolio current;
            PriceProfilePortfolio updated = TryLoad(workbook, out current)
                ? current.WithProfile(profile)
                : PriceProfilePortfolio.Create(new[] { profile });
            return Save(workbook, updated);
        }
    }
}
