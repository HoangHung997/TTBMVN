using System;
using System.IO;
using ExcelAddIn1.Core;

namespace ExcelAddIn1.LicenseTool
{
    internal static class LicensePrivateKeyStore
    {
        private const string SigningKeyPathEnvironmentVariable = "TTBMVN_LICENSE_SIGNING_KEY";

        public static string GetKeyPath()
        {
            string configured = Environment.GetEnvironmentVariable(SigningKeyPathEnvironmentVariable);
            if (!string.IsNullOrWhiteSpace(configured))
                return Path.GetFullPath(configured.Trim());

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TTBMVN Excel Tools",
                "Publisher",
                ProductKeyCodec.SigningKeyFileName);
        }

        public static byte[] Load()
        {
            string path = GetKeyPath();
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(
                    "Khong tim thay private key phat hanh. Xem docs/LICENSE_TOOL.md de cau hinh.",
                    path);
            }

            byte[] key = File.ReadAllBytes(path);
            if (key.Length == 0)
                throw new InvalidDataException("Private key phat hanh dang rong.");
            return key;
        }
    }
}
