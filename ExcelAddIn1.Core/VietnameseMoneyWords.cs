using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelAddIn1.Core
{
    public static class VietnameseMoneyWords
    {
        private static readonly string[] Digits =
        {
            "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"
        };

        private static readonly string[] Scales =
        {
            string.Empty, "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"
        };

        public static string ToWords(long value)
        {
            if (value < 0)
                throw new ArgumentOutOfRangeException(nameof(value));
            if (value == 0)
                return "Không đồng chẵn.";

            var groups = new List<int>();
            long remaining = value;
            while (remaining > 0)
            {
                groups.Add((int)(remaining % 1000));
                remaining /= 1000;
            }
            if (groups.Count > Scales.Length)
                throw new ArgumentOutOfRangeException(nameof(value), "Số tiền vượt phạm vi đọc hỗ trợ.");

            var parts = new List<string>();
            for (int index = groups.Count - 1; index >= 0; index--)
            {
                int group = groups[index];
                if (group == 0)
                    continue;
                bool full = index < groups.Count - 1 && group < 100;
                string text = ReadGroup(group, full);
                if (Scales[index].Length > 0)
                    text += " " + Scales[index];
                parts.Add(text);
            }

            string result = string.Join(", ", parts);
            return char.ToUpperInvariant(result[0]) + result.Substring(1) + " đồng chẵn.";
        }

        private static string ReadGroup(int value, bool full)
        {
            int hundreds = value / 100;
            int tens = (value % 100) / 10;
            int units = value % 10;
            var builder = new StringBuilder();

            if (hundreds > 0 || full)
            {
                builder.Append(Digits[hundreds]).Append(" trăm");
                if (tens == 0 && units > 0)
                    builder.Append(" linh");
            }
            if (tens > 1)
            {
                Append(builder, Digits[tens] + " mươi");
            }
            else if (tens == 1)
            {
                Append(builder, "mười");
            }
            if (units > 0)
            {
                string unit = Digits[units];
                if (units == 1 && tens > 1)
                    unit = "mốt";
                else if (units == 5 && tens > 0)
                    unit = "lăm";
                Append(builder, unit);
            }
            return builder.ToString();
        }

        private static void Append(StringBuilder builder, string value)
        {
            if (builder.Length > 0)
                builder.Append(' ');
            builder.Append(value);
        }
    }
}
