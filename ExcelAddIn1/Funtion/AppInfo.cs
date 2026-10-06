using System;
using System.IO;
using System.Reflection;

namespace ExcelAddIn1.Funtion
{
    public static class AppInfo
    {
        public const string ProductName = "TTBMVN Excel Tools";
        public const string CompanyName = "TTBMVN";
        public const string SafeProductFolder = "TTBMVNExcelTools";

        public static string Version
        {
            get
            {
                Version version = Assembly.GetExecutingAssembly().GetName().Version;
                return version == null ? "0.0.0.0" : version.ToString();
            }
        }
    }

    public static class AppPaths
    {
        public static string LocalAppDataRoot
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    AppInfo.SafeProductFolder);
            }
        }

        public static string LogDirectory => Path.Combine(LocalAppDataRoot, "Logs");

        public static string LogFilePath => Path.Combine(LogDirectory, "runtime.log");

        public static string SettingsBackupDirectory => Path.Combine(LocalAppDataRoot, "SettingsBackups");

        public static string RegulationPackageDirectory => Path.Combine(LocalAppDataRoot, "RegulationPackages");

        public static string PriceProfileDirectory => Path.Combine(LocalAppDataRoot, "PriceProfiles");

        public static string SupportDirectory => Path.Combine(LocalAppDataRoot, "SupportPackages");

        public static string UpdateDirectory => Path.Combine(LocalAppDataRoot, "Updates");

        public static string UpdateCenterStatePath => Path.Combine(LocalAppDataRoot, "update-center.state");

        public static string LicenseFilePath => Path.Combine(LocalAppDataRoot, "license.dat");

        public static string TrialFilePath => Path.Combine(LocalAppDataRoot, "trial.dat");

        public static void EnsureRoot()
        {
            Directory.CreateDirectory(LocalAppDataRoot);
        }
    }
}
