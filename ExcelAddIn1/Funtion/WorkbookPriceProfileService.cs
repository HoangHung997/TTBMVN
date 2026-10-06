using ExcelAddIn1.Core;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookPriceProfileService
    {
        public const string ManifestPropertyName = "TTBMVN.PriceProfile.Manifest";
        private const string PartPropertyPrefix = "TTBMVN.PriceProfile.Part.";
        private const string Label = "PriceProfile";

        public static bool Save(Excel.Workbook workbook, PriceProfile profile)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            ProjectProfile project = WorkbookProjectProfileService.LoadRequired(workbook);
            if (!string.Equals(project.PriceProfileId, profile.ProfileId, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "PriceProfileId khong khop ProjectProfile; hay pin profile truoc khi luu.");
            }
            return SavePayload(workbook, PriceProfileSerializer.Serialize(profile));
        }

        public static bool SaveAndPin(Excel.Workbook workbook, PriceProfile profile)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            PriceProfileValidationResult validation = PriceProfileValidator.Validate(profile);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(profile));

            string oldProjectPayload;
            if (!WorkbookProjectProfileService.TryReadPayload(workbook, out oldProjectPayload))
                throw new InvalidOperationException("Workbook chua co ProjectProfile.");
            string oldPricePayload;
            bool hadPrice = TryReadPayload(workbook, out oldPricePayload);
            ProjectProfile pinned = WorkbookProjectProfileService.LoadRequired(workbook).Clone();
            pinned.PriceProfileId = profile.ProfileId;

            try
            {
                using (new ExcelWriteContext(workbook.Application))
                {
                    WorkbookProjectProfileService.Save(workbook, pinned);
                    SavePayload(workbook, PriceProfileSerializer.Serialize(profile));
                }
                return true;
            }
            catch
            {
                try
                {
                    using (new ExcelWriteContext(workbook.Application))
                    {
                        WorkbookProjectProfileService.RestorePayload(workbook, oldProjectPayload);
                        if (hadPrice)
                            RestorePayload(workbook, oldPricePayload);
                        else
                            Clear(workbook);
                    }
                }
                catch (Exception rollbackError)
                {
                    RuntimeLogger.Log(rollbackError, "Rollback PriceProfile pin");
                }
                throw;
            }
        }

        public static bool TryLoad(Excel.Workbook workbook, out PriceProfile profile)
        {
            string payload;
            if (!TryReadPayload(workbook, out payload))
            {
                profile = null;
                return false;
            }
            profile = PriceProfileSerializer.Deserialize(payload);
            ProjectProfile project = WorkbookProjectProfileService.LoadRequired(workbook);
            if (!string.Equals(project.PriceProfileId, profile.ProfileId, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("PriceProfile snapshot khong khop PriceProfileId cua project.");
            return true;
        }

        public static PriceProfile LoadRequired(Excel.Workbook workbook)
        {
            PriceProfile profile;
            if (!TryLoad(workbook, out profile))
                throw new InvalidOperationException("Workbook chua co PriceProfile snapshot.");
            return profile;
        }

        public static bool RestorePayload(Excel.Workbook workbook, string payload)
        {
            if (payload == null)
                return Clear(workbook);
            PriceProfileSerializer.Deserialize(payload);
            return SavePayload(workbook, payload);
        }

        public static bool TryReadPayload(Excel.Workbook workbook, out string payload)
        {
            return WorkbookCustomPayloadStore.TryRead(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                Label,
                out payload);
        }

        public static bool Clear(Excel.Workbook workbook)
        {
            return WorkbookCustomPayloadStore.Clear(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                Label);
        }

        private static bool SavePayload(Excel.Workbook workbook, string payload)
        {
            return WorkbookCustomPayloadStore.Save(
                workbook,
                ManifestPropertyName,
                PartPropertyPrefix,
                payload,
                Label);
        }
    }
}
