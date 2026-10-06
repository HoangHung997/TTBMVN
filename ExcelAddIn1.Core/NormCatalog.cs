using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum NormResourceKind
    {
        Material = 1,
        Labor = 2,
        Machine = 3
    }

    public enum NormAdjustmentOperation
    {
        Factor = 1,
        Add = 2
    }

    public sealed class NormResourceRate
    {
        internal NormResourceRate(
            NormResourceKind kind,
            string resourceCode,
            string unit,
            IEnumerable<decimal> quantities)
        {
            Kind = kind;
            ResourceCode = resourceCode;
            Unit = unit;
            Quantities = new ReadOnlyCollection<decimal>(quantities.ToList());
        }

        public NormResourceKind Kind { get; }
        public string ResourceCode { get; }
        public string Unit { get; }
        public IReadOnlyList<decimal> Quantities { get; }
    }

    public sealed class NormAdjustment
    {
        internal NormAdjustment(
            NormAdjustmentOperation operation,
            string condition,
            IEnumerable<NormResourceKind> targetKinds,
            string resourceCode,
            string unit,
            decimal value)
        {
            Operation = operation;
            Condition = condition;
            TargetKinds = new ReadOnlyCollection<NormResourceKind>(targetKinds.ToList());
            ResourceCode = resourceCode;
            Unit = unit;
            Value = value;
        }

        public NormAdjustmentOperation Operation { get; }
        public string Condition { get; }
        public IReadOnlyList<NormResourceKind> TargetKinds { get; }
        public string ResourceCode { get; }
        public string Unit { get; }
        public decimal Value { get; }
    }

    public sealed class NormConstraint
    {
        internal NormConstraint(string operation, string argument1, string argument2)
        {
            Operation = operation;
            Argument1 = argument1;
            Argument2 = argument2;
        }

        public string Operation { get; }
        public string Argument1 { get; }
        public string Argument2 { get; }
    }

    public sealed class NormDefinition
    {
        internal NormDefinition(
            string key,
            string title,
            string workUnit,
            IEnumerable<string> variants,
            IEnumerable<NormResourceRate> rates,
            IEnumerable<NormAdjustment> adjustments,
            IEnumerable<NormConstraint> constraints,
            RegulationSourceLocator source)
        {
            Key = key;
            Title = title;
            WorkUnit = workUnit;
            Variants = new ReadOnlyCollection<string>(variants.ToList());
            Rates = new ReadOnlyCollection<NormResourceRate>(rates.ToList());
            Adjustments = new ReadOnlyCollection<NormAdjustment>(adjustments.ToList());
            Constraints = new ReadOnlyCollection<NormConstraint>(constraints.ToList());
            Source = source;
        }

        public string Key { get; }
        public string Title { get; }
        public string WorkUnit { get; }
        public IReadOnlyList<string> Variants { get; }
        public IReadOnlyList<NormResourceRate> Rates { get; }
        public IReadOnlyList<NormAdjustment> Adjustments { get; }
        public IReadOnlyList<NormConstraint> Constraints { get; }
        public RegulationSourceLocator Source { get; }
    }

    public sealed class NormCatalog
    {
        private readonly IReadOnlyDictionary<string, NormDefinition> byKey;

        private NormCatalog(IEnumerable<NormDefinition> definitions)
        {
            NormDefinition[] ordered = definitions
                .OrderBy(item => item.Key, StringComparer.Ordinal)
                .ToArray();
            Definitions = new ReadOnlyCollection<NormDefinition>(ordered);
            byKey = new ReadOnlyDictionary<string, NormDefinition>(
                ordered.ToDictionary(item => item.Key, StringComparer.Ordinal));
        }

        public IReadOnlyList<NormDefinition> Definitions { get; }

        public static NormCatalog Load(RegulationDataModule module)
        {
            if (module == null)
                throw new ArgumentNullException(nameof(module));
            if (module.Kind != RegulationModuleKind.Norm)
                throw new ArgumentException("Module khong phai Norm.", nameof(module));

            RegulationDataValidationResult validation = RegulationDataValidator.Validate(module);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(module));

            var definitions = new List<NormDefinition>();
            foreach (RegulationDataRecord record in module.Records)
            {
                if (string.Equals(record.RecordType, "NormCatalog", StringComparison.Ordinal))
                    definitions.Add(NormRecordParser.Parse(record));
                else if (!string.Equals(record.RecordType, "ProvisionalEstimateRate", StringComparison.Ordinal))
                    throw new FormatException("Norm co record type khong duoc ho tro: " + record.RecordType + ".");
            }
            if (definitions.Count == 0)
                throw new FormatException("Module Norm khong co NormCatalog.");
            return new NormCatalog(definitions);
        }

        public NormDefinition FindRequired(string key)
        {
            if (!byKey.TryGetValue((key ?? string.Empty).Trim(), out NormDefinition value))
                throw new KeyNotFoundException("Khong co dinh muc theo key: " + key + ".");
            return value;
        }
    }

    internal static class NormRecordParser
    {
        private static readonly ISet<string> AllowedFields = new HashSet<string>(
            new[] { "variantCodes", "rates", "adjustments", "constraints" },
            StringComparer.Ordinal);

        internal static NormDefinition Parse(RegulationDataRecord record)
        {
            Dictionary<string, string> fields = ParseFields(record);
            foreach (string field in fields.Keys)
            {
                if (!AllowedFields.Contains(field))
                    throw Error(record, "field khong duoc ho tro: " + field + ".");
            }

            string[] variants = SplitRequired(fields, record, "variantCodes", ',');
            EnsureUnique(record, variants, "variantCodes");
            string[] rateTexts = SplitRequired(fields, record, "rates", '|');
            var rates = new List<NormResourceRate>();
            var resourceKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (string rateText in rateTexts)
            {
                string[] parts = rateText.Split(':');
                if (parts.Length != 4)
                    throw Error(record, "rate phai co 4 thanh phan: " + rateText + ".");
                NormResourceKind kind = ParseKind(record, parts[0]);
                string resource = RequiredToken(record, parts[1], "resourceCode");
                string unit = RequiredToken(record, parts[2], "resourceUnit");
                string identity = kind + ":" + resource;
                if (!resourceKeys.Add(identity))
                    throw Error(record, "resource trung: " + identity + ".");
                string[] quantityTexts = parts[3].Split(',');
                if (quantityTexts.Length != variants.Length)
                    throw Error(record, "so quantity khong khop variant cho " + resource + ".");
                rates.Add(new NormResourceRate(
                    kind,
                    resource,
                    unit,
                    quantityTexts.Select(value => ParseNonNegative(record, value, resource))));
            }

            var adjustments = new List<NormAdjustment>();
            if (fields.TryGetValue("adjustments", out string adjustmentField))
            {
                foreach (string adjustmentText in Split(adjustmentField, '|'))
                    adjustments.Add(ParseAdjustment(record, adjustmentText));
            }

            var constraints = new List<NormConstraint>();
            if (fields.TryGetValue("constraints", out string constraintField))
            {
                foreach (string constraintText in Split(constraintField, '|'))
                    constraints.Add(ParseConstraint(record, constraintText));
            }

            return new NormDefinition(
                record.Key,
                record.Title,
                record.Unit,
                variants,
                rates,
                adjustments,
                constraints,
                record.Source);
        }

        private static NormAdjustment ParseAdjustment(RegulationDataRecord record, string text)
        {
            string[] parts = text.Split(':');
            if (parts.Length != 6)
                throw Error(record, "adjustment phai co 6 thanh phan: " + text + ".");
            NormAdjustmentOperation operation;
            if (parts[0] == "factor")
                operation = NormAdjustmentOperation.Factor;
            else if (parts[0] == "add")
                operation = NormAdjustmentOperation.Add;
            else
                throw Error(record, "adjustment operation khong hop le: " + parts[0] + ".");

            string condition = RequiredToken(record, parts[1], "condition");
            NormResourceKind[] targetKinds = Split(parts[2], '+')
                .Select(kindText => ParseKind(record, kindText))
                .Distinct()
                .ToArray();
            string resource = RequiredToken(record, parts[3], "adjustment resource");
            string unit = RequiredToken(record, parts[4], "adjustment unit");
            decimal value = ParseNonNegative(record, parts[5], "adjustment value");
            if (operation == NormAdjustmentOperation.Factor)
            {
                if (resource != "*" || unit != "ratio" || value <= 0m)
                    throw Error(record, "factor phai dung resource=* va unit=ratio, value > 0.");
            }
            else if (resource == "*" || unit == "ratio")
            {
                throw Error(record, "add phai chi dinh resource va unit cu the.");
            }
            return new NormAdjustment(operation, condition, targetKinds, resource, unit, value);
        }

        private static NormConstraint ParseConstraint(RegulationDataRecord record, string text)
        {
            string[] parts = text.Split(':');
            if (parts.Length == 2 &&
                (parts[0] == "prohibit-condition" || parts[0] == "external-estimate"))
            {
                return new NormConstraint(parts[0], RequiredToken(record, parts[1], "constraint"), string.Empty);
            }
            if (parts.Length == 3 && parts[0] == "requires-condition")
            {
                return new NormConstraint(
                    parts[0],
                    RequiredToken(record, parts[1], "constraint variant"),
                    RequiredToken(record, parts[2], "constraint condition"));
            }
            throw Error(record, "constraint khong hop le: " + text + ".");
        }

        private static Dictionary<string, string> ParseFields(RegulationDataRecord record)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string item in Split(record.Data, ';'))
            {
                int separator = item.IndexOf('=');
                if (separator <= 0 || separator == item.Length - 1)
                    throw Error(record, "data field khong hop le: " + item + ".");
                string key = item.Substring(0, separator).Trim();
                string value = item.Substring(separator + 1).Trim();
                if (result.ContainsKey(key))
                    throw Error(record, "data field trung: " + key + ".");
                result.Add(key, value);
            }
            return result;
        }

        private static string[] SplitRequired(
            Dictionary<string, string> fields,
            RegulationDataRecord record,
            string field,
            char separator)
        {
            if (!fields.TryGetValue(field, out string value))
                throw Error(record, "thieu field: " + field + ".");
            return Split(value, separator);
        }

        private static string[] Split(string value, char separator)
        {
            return (value ?? string.Empty)
                .Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => item.Length > 0)
                .ToArray();
        }

        private static NormResourceKind ParseKind(RegulationDataRecord record, string value)
        {
            if (!Enum.TryParse(value, false, out NormResourceKind kind) ||
                !Enum.IsDefined(typeof(NormResourceKind), kind))
            {
                throw Error(record, "resource kind khong hop le: " + value + ".");
            }
            return kind;
        }

        private static decimal ParseNonNegative(
            RegulationDataRecord record,
            string value,
            string field)
        {
            if (value.Contains(",") ||
                !decimal.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal parsed) ||
                parsed < 0m)
            {
                throw Error(record, field + " phai la so invariant khong am: " + value + ".");
            }
            return parsed;
        }

        private static string RequiredToken(RegulationDataRecord record, string value, string field)
        {
            string parsed = (value ?? string.Empty).Trim();
            if (parsed.Length == 0)
                throw Error(record, field + " khong duoc trong.");
            return parsed;
        }

        private static void EnsureUnique(RegulationDataRecord record, string[] values, string field)
        {
            if (values.Distinct(StringComparer.Ordinal).Count() != values.Length)
                throw Error(record, field + " co gia tri trung.");
        }

        private static FormatException Error(RegulationDataRecord record, string message)
        {
            return new FormatException(record.Key + ": " + message);
        }
    }
}
