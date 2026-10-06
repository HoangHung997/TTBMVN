using System;
using System.Globalization;
using System.IO;
using System.Text;
using ExcelAddIn1.Core;

namespace ExcelAddIn1.Funtion
{
    public sealed class LicenseStatus
    {
        public bool IsActivated { get; set; }
        public bool IsTrial { get; set; }
        public bool IsExpired { get; set; }
        public DateTime? ExpiryDate { get; set; }
        public int TrialDaysRemaining { get; set; }
        public string MachineId { get; set; }
        public string Message { get; set; }
    }

    public static class LicenseManager
    {
        private const int TrialDays = 30;

        public static LicenseStatus GetStatus()
        {
            string machineId = GetMachineId();
            string licenseKey = ReadFileSafe(AppPaths.LicenseFilePath);
            if (!string.IsNullOrWhiteSpace(licenseKey) && TryValidateLicenseKey(licenseKey.Trim(), machineId, out DateTime expiryDate))
            {
                bool expired = DateTime.Today > expiryDate.Date;
                return new LicenseStatus
                {
                    IsActivated = !expired,
                    IsExpired = expired,
                    ExpiryDate = expiryDate.Date,
                    MachineId = machineId,
                    Message = expired
                        ? "License da het han."
                        : "License dang hoat dong den " + expiryDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + "."
                };
            }

            DateTime trialStart = GetOrCreateTrialStartDate();
            int usedDays = Math.Max(0, (DateTime.Today - trialStart.Date).Days);
            int remaining = Math.Max(0, TrialDays - usedDays);
            bool trialExpired = remaining <= 0;
            return new LicenseStatus
            {
                IsTrial = !trialExpired,
                IsExpired = trialExpired,
                TrialDaysRemaining = remaining,
                MachineId = machineId,
                Message = trialExpired
                    ? "Ban dung thu da het han."
                    : "Ban dung thu con " + remaining.ToString(CultureInfo.InvariantCulture) + " ngay."
            };
        }

        public static bool TryActivate(string licenseKey, out string message)
        {
            string machineId = GetMachineId();
            if (!TryValidateLicenseKey(licenseKey, machineId, out DateTime expiryDate))
            {
                message = "License key khong hop le cho may nay.";
                return false;
            }

            Directory.CreateDirectory(AppPaths.LocalAppDataRoot);
            File.WriteAllText(AppPaths.LicenseFilePath, licenseKey.Trim(), Encoding.UTF8);
            message = "Da kich hoat license den " + expiryDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + ".";
            return true;
        }

        public static string GetMachineId()
        {
            return HardwareFingerprintProvider.GetMachineId();
        }

        private static DateTime GetOrCreateTrialStartDate()
        {
            string text = ReadFileSafe(AppPaths.TrialFilePath);
            if (DateTime.TryParseExact(text, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                return date.Date;

            Directory.CreateDirectory(AppPaths.LocalAppDataRoot);
            date = DateTime.Today;
            File.WriteAllText(AppPaths.TrialFilePath, date.ToString("yyyyMMdd", CultureInfo.InvariantCulture), Encoding.UTF8);
            return date;
        }

        private static bool TryValidateLicenseKey(string licenseKey, string machineId, out DateTime expiryDate)
        {
            expiryDate = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(licenseKey))
                return false;

            return ProductKeyCodec.TryValidate(licenseKey, machineId, out expiryDate);
        }

        private static string ReadFileSafe(string path)
        {
            try
            {
                return File.Exists(path) ? File.ReadAllText(path, Encoding.UTF8).Trim() : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
