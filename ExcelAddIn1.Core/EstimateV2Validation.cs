using System;
using System.Globalization;
using System.Text;

namespace ExcelAddIn1.Core
{
    public static class EstimateV2ValidationRules
    {
        public static string[] ExpectedCostRowFormulas(
            int quantityColumn,
            int row,
            string rateId)
        {
            if (quantityColumn < 1 || quantityColumn > 16378)
                throw new ArgumentOutOfRangeException(nameof(quantityColumn));
            if (row < 1 || row > 1048576)
                throw new ArgumentOutOfRangeException(nameof(row));

            string quantity =
                ExcelColumnAddress.ToLetters(quantityColumn) +
                row.ToString(CultureInfo.InvariantCulture);

            var output = new string[6];
            output[0] = "=" +
                EstimateV2ExcelNames.RateComponent(rateId, "VL");
            output[1] = "=" +
                EstimateV2ExcelNames.RateComponent(rateId, "NC");
            output[2] = "=" +
                EstimateV2ExcelNames.RateComponent(rateId, "M");

            for (int index = 0; index < 3; index++)
            {
                string unitPrice =
                    ExcelColumnAddress.ToLetters(
                        quantityColumn + 1 + index) +
                    row.ToString(CultureInfo.InvariantCulture);
                output[3 + index] =
                    "=" + quantity + "*" + unitPrice;
            }

            return output;
        }

        public static string[] ExpectedThkpDirectFormulas()
        {
            return new[]
            {
                "=" + EstimateV2ExcelNames.EstimateTotal("VL"),
                "=" + EstimateV2ExcelNames.EstimateTotal("NC"),
                "=" + EstimateV2ExcelNames.EstimateTotal("M"),
                "=" + EstimateV2ExcelNames.EstimateTotal("TOTAL")
            };
        }

        public static bool FormulaEquivalent(
            string actual,
            string expected)
        {
            return string.Equals(
                NormalizeFormula(actual),
                NormalizeFormula(expected),
                StringComparison.OrdinalIgnoreCase);
        }

        public static string NormalizeFormula(string formula)
        {
            string value = (formula ?? string.Empty).Trim();
            if (value.Length == 0)
                return string.Empty;

            var builder = new StringBuilder(value.Length);
            bool inString = false;
            for (int index = 0; index < value.Length; index++)
            {
                char ch = value[index];
                if (ch == '"')
                {
                    inString = !inString;
                    builder.Append(ch);
                    continue;
                }

                if (!inString)
                {
                    if (char.IsWhiteSpace(ch) || ch == '$')
                        continue;
                    if (ch == ';')
                    {
                        builder.Append(',');
                        continue;
                    }
                }

                builder.Append(ch);
            }

            string normalized = builder.ToString();
            if (normalized.StartsWith(
                "_xlfn.",
                StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(6);
            }
            else if (normalized.StartsWith(
                "=_xlfn.",
                StringComparison.OrdinalIgnoreCase))
            {
                normalized = "=" + normalized.Substring(7);
            }

            return normalized.ToUpperInvariant();
        }

        public static string ExcelErrorText(int errorCode)
        {
            int code = errorCode & 0xFFFF;
            switch (code)
            {
                case 2000:
                    return "#NULL!";
                case 2007:
                    return "#DIV/0!";
                case 2015:
                    return "#VALUE!";
                case 2023:
                    return "#REF!";
                case 2029:
                    return "#NAME?";
                case 2036:
                    return "#NUM!";
                case 2042:
                    return "#N/A";
                case 2043:
                    return "#GETTING_DATA";
                default:
                    return "#ERROR!";
            }
        }
    }
}
