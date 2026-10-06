using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace ExcelAddIn1.Core
{
    public sealed class WorkbookSheetDescriptor
    {
        public WorkbookSheetDescriptor(string key, string name)
        {
            Key = (key ?? string.Empty).Trim();
            Name = (name ?? string.Empty).Trim();
            if (Key.Length == 0)
                throw new ArgumentException("Sheet key khong duoc trong.", nameof(key));
            if (Name.Length == 0)
                throw new ArgumentException("Sheet name khong duoc trong.", nameof(name));
        }

        public string Key { get; }
        public string Name { get; }

        public override string ToString()
        {
            return Name;
        }
    }

    public static class WorkbookSheetList
    {
        public static IReadOnlyList<WorkbookSheetDescriptor> Copy(
            IEnumerable<WorkbookSheetDescriptor> sheets)
        {
            return new ReadOnlyCollection<WorkbookSheetDescriptor>(
                (sheets ?? Enumerable.Empty<WorkbookSheetDescriptor>())
                    .Where(sheet => sheet != null)
                    .ToList());
        }

        public static string BuildSignature(IEnumerable<WorkbookSheetDescriptor> sheets)
        {
            var builder = new StringBuilder();
            foreach (WorkbookSheetDescriptor sheet in sheets ?? Enumerable.Empty<WorkbookSheetDescriptor>())
            {
                if (sheet == null)
                    continue;

                AppendLengthPrefixed(builder, sheet.Key);
                AppendLengthPrefixed(builder, sheet.Name);
            }
            return builder.ToString();
        }

        public static WorkbookSheetDescriptor ResolveSelection(
            IEnumerable<WorkbookSheetDescriptor> sheets,
            string previousKey,
            string previousName)
        {
            IReadOnlyList<WorkbookSheetDescriptor> list = Copy(sheets);
            if (!string.IsNullOrWhiteSpace(previousKey))
            {
                WorkbookSheetDescriptor byKey = list.FirstOrDefault(sheet =>
                    string.Equals(sheet.Key, previousKey.Trim(), StringComparison.OrdinalIgnoreCase));
                if (byKey != null)
                    return byKey;
            }

            if (!string.IsNullOrWhiteSpace(previousName))
            {
                WorkbookSheetDescriptor byName = list.FirstOrDefault(sheet =>
                    string.Equals(sheet.Name, previousName.Trim(), StringComparison.OrdinalIgnoreCase));
                if (byName != null)
                    return byName;
            }

            return list.Count > 0 ? list[0] : null;
        }

        private static void AppendLengthPrefixed(StringBuilder builder, string value)
        {
            string text = value ?? string.Empty;
            builder.Append(text.Length);
            builder.Append(':');
            builder.Append(text);
            builder.Append('|');
        }
    }
}

