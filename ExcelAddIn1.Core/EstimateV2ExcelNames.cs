using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAddIn1.Core
{
    public static class EstimateV2ExcelNames
    {
        public static string ResourcePrice(
            NormResourceKind kind,
            string resourceCode,
            string unit,
            string packageIdentity = null)
        {
            string canonical = string.Join("|", new[]
            {
                ((int)kind).ToString(CultureInfo.InvariantCulture),
                (resourceCode ?? string.Empty).Trim().ToUpperInvariant(),
                (unit ?? string.Empty).Trim().ToUpperInvariant(),
                (packageIdentity ?? string.Empty).Trim().ToUpperInvariant()
            });
            return "TTBMVN_V2_PRICE_" + Hash(canonical).Substring(0, 20);
        }

        public static string LaborInput(string laborCode, string inputKind)
        {
            string canonical = "LABOR|" +
                (laborCode ?? string.Empty).Trim().ToUpperInvariant() + "|" +
                (inputKind ?? string.Empty).Trim().ToUpperInvariant();
            return "TTBMVN_V2_INPUT_" + Hash(canonical).Substring(0, 20);
        }

        public static string EstimateTotal(string component)
        {
            string normalized =
                (component ?? string.Empty).Trim().ToUpperInvariant();
            switch (normalized)
            {
                case "VL":
                    return "TTBMVN_V2_GIADT_VL";
                case "NC":
                    return "TTBMVN_V2_GIADT_NC";
                case "M":
                    return "TTBMVN_V2_GIADT_M";
                case "TOTAL":
                    return "TTBMVN_V2_GIADT_TOTAL";
                default:
                    throw new ArgumentException(
                        "Component tong hop khong hop le: " + component,
                        nameof(component));
            }
        }

        public static string RateComponent(
            string rateId,
            string component)
        {
            string canonical = "RATE|" +
                (rateId ?? string.Empty).Trim().ToUpperInvariant() + "|" +
                (component ?? string.Empty).Trim().ToUpperInvariant();
            return "TTBMVN_V2_RATE_" + Hash(canonical).Substring(0, 20);
        }

        public static string FuelInput(string priceCode)
        {
            string canonical = "FUEL|" +
                (priceCode ?? string.Empty).Trim().ToUpperInvariant();
            return "TTBMVN_V2_INPUT_" + Hash(canonical).Substring(0, 20);
        }

        private static string Hash(string value)
        {
            using (SHA256 algorithm = SHA256.Create())
            {
                byte[] bytes = algorithm.ComputeHash(
                    Encoding.UTF8.GetBytes(value ?? string.Empty));
                var builder = new StringBuilder(bytes.Length * 2);
                foreach (byte item in bytes)
                    builder.Append(item.ToString("X2", CultureInfo.InvariantCulture));
                return builder.ToString();
            }
        }
    }
}
