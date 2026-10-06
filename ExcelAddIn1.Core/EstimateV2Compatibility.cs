using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ExcelAddIn1.Core
{
    public enum EstimateV2LegacySheetKind
    {
        Unknown = 0,
        EstimateAppendix = 1,
        CostSummary = 2,
        ResourcePrices = 3,
        UnitRateLand = 4,
        UnitRateWater = 5,
        UnitRateSea = 6
    }

    public enum EstimateV2LegacyAudienceHint
    {
        None = 0,
        StateBudgetSalary = 1,
        NonStateSalary = 2
    }

    public sealed class EstimateV2LegacySheetClassification
    {
        internal EstimateV2LegacySheetClassification(
            EstimateV2LegacySheetKind kind,
            EstimateV2LegacyAudienceHint audienceHint,
            bool isCanonicalName,
            string normalizedName)
        {
            Kind = kind;
            AudienceHint = audienceHint;
            IsCanonicalName = isCanonicalName;
            NormalizedName = normalizedName ?? string.Empty;
        }

        public EstimateV2LegacySheetKind Kind { get; }
        public EstimateV2LegacyAudienceHint AudienceHint { get; }
        public bool IsCanonicalName { get; }
        public string NormalizedName { get; }
        public bool IsKnown => Kind != EstimateV2LegacySheetKind.Unknown;
        public bool IsLegacyAlias => IsKnown && !IsCanonicalName;
    }

    public static class EstimateV2CompatibilityRules
    {
        public static EstimateV2LegacySheetClassification ClassifySheetName(
            string sheetName)
        {
            string normalized = Normalize(sheetName);
            EstimateV2LegacySheetKind kind =
                EstimateV2LegacySheetKind.Unknown;
            string canonical = string.Empty;

            if (StartsWith(normalized, "THKP TC"))
            {
                kind = EstimateV2LegacySheetKind.CostSummary;
                canonical = "THKP TC";
            }
            else if (StartsWith(normalized, "GIA DT TC"))
            {
                kind = EstimateV2LegacySheetKind.EstimateAppendix;
                canonical = "GIA DT TC";
            }
            else if (StartsWith(normalized, "DG CAN"))
            {
                kind = EstimateV2LegacySheetKind.UnitRateLand;
                canonical = "DG CAN";
            }
            else if (StartsWith(normalized, "DG NUOC"))
            {
                kind = EstimateV2LegacySheetKind.UnitRateWater;
                canonical = "DG NUOC";
            }
            else if (StartsWith(normalized, "DG BIEN"))
            {
                kind = EstimateV2LegacySheetKind.UnitRateSea;
                canonical = "DG BIEN";
            }
            else if (StartsWith(normalized, "VL NC M"))
            {
                kind = EstimateV2LegacySheetKind.ResourcePrices;
                canonical = "VL NC M";
            }

            EstimateV2LegacyAudienceHint audience =
                EstimateV2LegacyAudienceHint.None;
            string[] tokens = normalized.Split(
                new[] { ' ' },
                StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Contains("VT") ||
                normalized.Contains("HUONG LUONG NSNN") ||
                normalized.Contains("NGAN SACH"))
            {
                audience =
                    EstimateV2LegacyAudienceHint.StateBudgetSalary;
            }
            else if (tokens.Contains("DN") ||
                normalized.Contains("DOANH NGHIEP"))
            {
                audience =
                    EstimateV2LegacyAudienceHint.NonStateSalary;
            }

            bool canonicalName =
                kind != EstimateV2LegacySheetKind.Unknown &&
                string.Equals(
                    normalized,
                    canonical,
                    StringComparison.OrdinalIgnoreCase);

            return new EstimateV2LegacySheetClassification(
                kind,
                audience,
                canonicalName,
                normalized);
        }

        public static bool LooksLikeV2EstimateHeader(
            IEnumerable<string> headerValues)
        {
            HashSet<string> values = NormalizeSet(headerValues);
            return values.Contains("MA CONG TAC") &&
                values.Contains("DINH MUC") &&
                HasDescription(values) &&
                HasUnit(values) &&
                HasQuantity(values);
        }

        public static bool LooksLikeLegacyEstimateHeader(
            IEnumerable<string> headerValues)
        {
            HashSet<string> values = NormalizeSet(headerValues);
            return values.Contains("TT") &&
                HasDescription(values) &&
                HasUnit(values) &&
                HasQuantity(values) &&
                (values.Any(value =>
                    value.StartsWith(
                        "DON GIA",
                        StringComparison.Ordinal)) ||
                 values.Any(value =>
                    value.StartsWith(
                        "THANH TIEN",
                        StringComparison.Ordinal)));
        }

        public static string Normalize(string value)
        {
            string decomposed =
                (value ?? string.Empty)
                    .Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder();
            foreach (char ch in decomposed)
            {
                UnicodeCategory category =
                    CharUnicodeInfo.GetUnicodeCategory(ch);
                if (category != UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(
                        char.ToUpperInvariant(ch));
                }
            }

            string text = builder.ToString()
                .Replace('Đ', 'D')
                .Replace('đ', 'D');

            return string.Join(
                " ",
                text.Split(
                    new[]
                    {
                        ' ', '\t', '\r', '\n',
                        '-', '_', '/', '(', ')',
                        '.', ':'
                    },
                    StringSplitOptions.RemoveEmptyEntries));
        }

        private static bool StartsWith(
            string value,
            string prefix)
        {
            return value.Equals(
                    prefix,
                    StringComparison.OrdinalIgnoreCase) ||
                value.StartsWith(
                    prefix + " ",
                    StringComparison.OrdinalIgnoreCase);
        }

        private static HashSet<string> NormalizeSet(
            IEnumerable<string> values)
        {
            return new HashSet<string>(
                (values ?? Enumerable.Empty<string>())
                    .Select(Normalize)
                    .Where(value => value.Length > 0),
                StringComparer.OrdinalIgnoreCase);
        }

        private static bool HasDescription(
            ISet<string> values)
        {
            return values.Contains("MO TA CONG VIEC") ||
                values.Contains("TEN CONG TAC") ||
                values.Contains("NOI DUNG CONG VIEC") ||
                values.Contains("NOI DUNG");
        }

        private static bool HasUnit(
            ISet<string> values)
        {
            return values.Contains("DON VI") ||
                values.Contains("DVT");
        }

        private static bool HasQuantity(
            ISet<string> values)
        {
            return values.Contains("KHOI LUONG") ||
                values.Contains("KL");
        }
    }
}
