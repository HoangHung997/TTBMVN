using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public static class RegulationRecordTable
    {
        public static IDictionary<string, string> ReadFields(string data)
        {
            var fields = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string item in (data ?? "").Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int index = item.IndexOf('=');
                if (index <= 0) throw new FormatException("Thong so phai co ten=gia tri.");
                fields.Add(item.Substring(0, index).Trim(), item.Substring(index + 1).Trim());
            }
            return fields;
        }

        public static string WriteFields(IEnumerable<KeyValuePair<string, string>> fields)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var field in fields)
            {
                if (string.IsNullOrWhiteSpace(field.Key) || field.Key.IndexOfAny(new[] { '=', ';', '\t', '\r', '\n' }) >= 0 ||
                    (field.Value ?? "").IndexOfAny(new[] { ';', '\t', '\r', '\n' }) >= 0)
                    throw new FormatException("Ten/gia tri thong so khong hop le.");
                result.Add(field.Key.Trim(), field.Value ?? "");
            }
            return string.Join(";", result.Select(p => p.Key + "=" + p.Value));
        }

        public static NormDefinition ReadNorm(RegulationDataRecord record) => NormRecordParser.Parse(record);

        public static string WriteNorm(string original, IEnumerable<string> variants, IEnumerable<string[]> rows)
        {
            var fields = ReadFields(original);
            string[] codes = variants.ToArray();
            if (codes.Length == 0 || codes.Any(c => string.IsNullOrWhiteSpace(c) || c.IndexOfAny(new[] { ',', ';', ':', '|', '\t', '\r', '\n' }) >= 0) ||
                codes.Distinct(StringComparer.Ordinal).Count() != codes.Length)
                throw new FormatException("Ma bien the rong, trung hoac co ky tu phan cach.");
            var rates = new List<string>();
            foreach (string[] row in rows)
            {
                if (row.Length != codes.Length + 3) throw new FormatException("So cot hao phi khong khop bien the.");
                if (row.Take(3).Any(s => string.IsNullOrWhiteSpace(s) || s.IndexOfAny(new[] { ':', '|', ';', ',', '\t', '\r', '\n' }) >= 0))
                    throw new FormatException("Ma hao phi/don vi khong hop le.");
                var quantities = row.Skip(3).Select(s => {
                    decimal value;
                    if (!decimal.TryParse(s, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value) || value < 0)
                        throw new FormatException("Hao phi phai la so khong am.");
                    return value.ToString(CultureInfo.InvariantCulture);
                });
                rates.Add(string.Join(":", row.Take(3)) + ":" + string.Join(",", quantities));
            }
            fields["variantCodes"] = string.Join(",", codes);
            fields["rates"] = string.Join("|", rates);
            string data = WriteFields(fields);
            NormRecordParser.Parse(new RegulationDataRecord("EDITOR", "NormCatalog", "unit", "Editor", data,
                new RegulationSourceLocator("EDITOR", 1, 1, "Editor"), RegulationDataVerification.Unverified));
            return data;
        }
    }
}
