using System;

namespace ExcelAddIn1.Core
{
    public static class ExcelColumnAddress
    {
        public static string Normalize(string column)
        {
            string text = (column ?? string.Empty).Trim().ToUpperInvariant();
            if (text.Length == 0)
                throw new ArgumentException("Cot khong hop le.");

            foreach (char c in text)
            {
                if (c < 'A' || c > 'Z')
                    throw new ArgumentException("Cot khong hop le: " + column);
            }

            return text;
        }

        public static bool IsValid(string column)
        {
            try
            {
                Normalize(column);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static int ToNumber(string column)
        {
            string text = Normalize(column);
            int result = 0;
            foreach (char c in text)
                result = result * 26 + (c - 'A' + 1);

            return result;
        }

        public static string ToLetters(int column)
        {
            if (column < 1 || column > 16384)
                throw new ArgumentOutOfRangeException(nameof(column));
            string result = string.Empty;
            int value = column;
            while (value > 0)
            {
                value--;
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }
            return result;
        }
    }
}
