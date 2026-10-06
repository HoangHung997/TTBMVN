using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum MachineRateAudience
    {
        StateBudgetSalary = 1,
        NonStateSalary = 2
    }

    public enum MachineFuelKind
    {
        None = 0,
        Battery = 1,
        Diesel = 2,
        Gasoline = 3,
        Electricity = 4,
        Other = 5
    }

    public sealed class MachineFuelRequirement
    {
        internal MachineFuelRequirement(
            decimal quantity,
            string unit,
            string priceCode,
            MachineFuelKind kind)
        {
            Quantity = quantity;
            Unit = unit;
            PriceCode = priceCode;
            Kind = kind;
        }

        public decimal Quantity { get; }
        public string Unit { get; }
        public string PriceCode { get; }
        public MachineFuelKind Kind { get; }
        public bool IsNone => Kind == MachineFuelKind.None;
    }

    public sealed class MachineOperatorRequirement
    {
        internal MachineOperatorRequirement(decimal quantity, string laborCode)
        {
            Quantity = quantity;
            LaborCode = laborCode;
        }

        public decimal Quantity { get; }
        public string LaborCode { get; }
    }

    public sealed class MachineRateDefinition
    {
        internal MachineRateDefinition(
            string key,
            string code,
            string title,
            MachineRateAudience audience,
            int annualShifts,
            decimal depreciationPercent,
            decimal repairPercent,
            decimal otherPercent,
            MachineFuelRequirement fuel,
            IEnumerable<MachineOperatorRequirement> operators,
            decimal referencePriceVnd,
            decimal referenceRecoverableValuePercent,
            RegulationSourceLocator source)
        {
            Key = key;
            Code = code;
            Title = title;
            Audience = audience;
            AnnualShifts = annualShifts;
            DepreciationPercent = depreciationPercent;
            RepairPercent = repairPercent;
            OtherPercent = otherPercent;
            Fuel = fuel;
            Operators = new ReadOnlyCollection<MachineOperatorRequirement>(operators.ToList());
            ReferencePriceVnd = referencePriceVnd;
            ReferenceRecoverableValuePercent = referenceRecoverableValuePercent;
            Source = source;
        }

        public string Key { get; }
        public string Code { get; }
        public string Title { get; }
        public MachineRateAudience Audience { get; }
        public int AnnualShifts { get; }
        public decimal DepreciationPercent { get; }
        public decimal RepairPercent { get; }
        public decimal OtherPercent { get; }
        public MachineFuelRequirement Fuel { get; }
        public IReadOnlyList<MachineOperatorRequirement> Operators { get; }
        public decimal ReferencePriceVnd { get; }
        public decimal ReferenceRecoverableValuePercent { get; }
        public RegulationSourceLocator Source { get; }
    }

    public sealed class MachineRateCatalog
    {
        private readonly IReadOnlyDictionary<string, MachineRateDefinition> byKey;
        private readonly IReadOnlyDictionary<string, MachineRateDefinition> byCode;

        private MachineRateCatalog(
            MachineRateAudience audience,
            IEnumerable<MachineRateDefinition> definitions)
        {
            Audience = audience;
            MachineRateDefinition[] ordered = definitions
                .OrderBy(item => item.Key, StringComparer.Ordinal)
                .ToArray();
            Definitions = new ReadOnlyCollection<MachineRateDefinition>(ordered);
            byKey = new ReadOnlyDictionary<string, MachineRateDefinition>(
                ordered.ToDictionary(item => item.Key, StringComparer.Ordinal));
            byCode = new ReadOnlyDictionary<string, MachineRateDefinition>(
                ordered.ToDictionary(item => item.Code, StringComparer.OrdinalIgnoreCase));
        }

        public MachineRateAudience Audience { get; }
        public IReadOnlyList<MachineRateDefinition> Definitions { get; }

        public static MachineRateCatalog Load(
            RegulationDataModule module,
            MachineRateAudience audience)
        {
            if (module == null)
                throw new ArgumentNullException(nameof(module));
            if (module.Kind != RegulationModuleKind.MachineRate)
                throw new ArgumentException("Module khong phai MachineRate.", nameof(module));
            if (!Enum.IsDefined(typeof(MachineRateAudience), audience))
                throw new ArgumentOutOfRangeException(nameof(audience));

            RegulationDataValidationResult validation = RegulationDataValidator.Validate(module);
            if (!validation.IsValid)
                throw new ArgumentException(string.Join(" ", validation.Errors), nameof(module));

            var definitions = new List<MachineRateDefinition>();
            foreach (RegulationDataRecord record in module.Records)
            {
                if (!string.Equals(record.RecordType, "MachineBaseData", StringComparison.Ordinal))
                    throw new FormatException("MachineRate co record type khong duoc ho tro: " + record.RecordType + ".");
                definitions.Add(MachineRateRecordParser.Parse(record, audience));
            }
            return new MachineRateCatalog(audience, definitions);
        }

        public MachineRateDefinition FindRequiredByKey(string key)
        {
            if (!byKey.TryGetValue((key ?? string.Empty).Trim(), out MachineRateDefinition value))
                throw new KeyNotFoundException("Khong co may theo key: " + key + ".");
            return value;
        }

        public MachineRateDefinition FindRequiredByCode(string code)
        {
            if (!byCode.TryGetValue((code ?? string.Empty).Trim(), out MachineRateDefinition value))
                throw new KeyNotFoundException("Khong co may theo ma: " + code + ".");
            return value;
        }
    }

    internal static class MachineRateRecordParser
    {
        private static readonly ISet<string> AllowedFields = new HashSet<string>(
            new[]
            {
                "annualShifts",
                "depreciationPercent",
                "repairPercent",
                "otherPercent",
                "fuel",
                "operators",
                "referencePriceVnd",
                "recoverableVatPercent",
                "stateCode",
                "nonStateCode",
                "nonStateFuel",
                "nonStateOperators",
                "nonStateReferencePriceVnd",
                "nonStateRecoverableValuePercent"
            },
            StringComparer.Ordinal);

        internal static MachineRateDefinition Parse(
            RegulationDataRecord record,
            MachineRateAudience audience)
        {
            Dictionary<string, string> fields = ParseFields(record.Data);
            foreach (string field in fields.Keys)
            {
                if (!AllowedFields.Contains(field))
                    throw Error(record, "field khong duoc ho tro: " + field + ".");
            }

            int annualShifts = ParseInt(record, Required(fields, record, "annualShifts"), "annualShifts");
            decimal depreciation = ParseNonNegativeDecimal(
                record, Required(fields, record, "depreciationPercent"), "depreciationPercent");
            decimal repair = ParseNonNegativeDecimal(
                record, Required(fields, record, "repairPercent"), "repairPercent");
            decimal other = ParseNonNegativeDecimal(
                record, Required(fields, record, "otherPercent"), "otherPercent");
            if (annualShifts <= 0)
                throw Error(record, "annualShifts phai lon hon 0.");
            if (depreciation + repair + other >= 100m)
                throw Error(record, "tong dinh muc phan tram phai nho hon 100.");

            bool nonState = audience == MachineRateAudience.NonStateSalary;
            string code = Select(
                fields,
                nonState ? "nonStateCode" : "stateCode",
                DeriveCode(record.Key, nonState));
            string fuelText = Select(
                fields,
                nonState ? "nonStateFuel" : string.Empty,
                Required(fields, record, "fuel"));
            string operatorsText = Select(
                fields,
                nonState ? "nonStateOperators" : string.Empty,
                Required(fields, record, "operators"));
            string priceText = Select(
                fields,
                nonState ? "nonStateReferencePriceVnd" : string.Empty,
                Required(fields, record, "referencePriceVnd"));
            string recoveryText = Select(
                fields,
                nonState ? "nonStateRecoverableValuePercent" : string.Empty,
                Required(fields, record, "recoverableVatPercent"));

            if (string.IsNullOrWhiteSpace(code))
                throw Error(record, "ma may trong.");
            decimal price = ParseNonNegativeDecimal(record, priceText, "referencePriceVnd");
            if (price <= 0m)
                throw Error(record, "referencePriceVnd phai lon hon 0.");
            decimal recovery = ParseNonNegativeDecimal(
                record, recoveryText, "recoverableValuePercent");
            if (recovery > 100m)
                throw Error(record, "recoverableValuePercent vuot 100.");

            decimal expectedRecovery = price >= MachineRateCalculator.RecoveryThresholdVnd
                ? MachineRateCalculator.StandardRecoveryPercent
                : 0m;
            if (recovery != expectedRecovery)
            {
                throw Error(
                    record,
                    "gia tri thu hoi khong khop nguong TT122 cho nguyen gia da chon.");
            }

            return new MachineRateDefinition(
                record.Key,
                code,
                record.Title,
                audience,
                annualShifts,
                depreciation,
                repair,
                other,
                ParseFuel(record, fuelText),
                ParseOperators(record, operatorsText),
                price,
                recovery,
                record.Source);
        }

        private static Dictionary<string, string> ParseFields(string value)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string part in (value ?? string.Empty).Split(';'))
            {
                int separator = part.IndexOf('=');
                if (separator <= 0 || separator == part.Length - 1)
                    throw new FormatException("MachineRate data khong dung key=value.");
                string key = part.Substring(0, separator);
                string item = part.Substring(separator + 1);
                if (result.ContainsKey(key))
                    throw new FormatException("MachineRate data trung field: " + key + ".");
                result.Add(key, item);
            }
            return result;
        }

        private static MachineFuelRequirement ParseFuel(RegulationDataRecord record, string value)
        {
            if (string.Equals(value, "none", StringComparison.Ordinal))
                return new MachineFuelRequirement(0m, string.Empty, string.Empty, MachineFuelKind.None);
            string[] parts = value.Split('-');
            if (parts.Length < 3)
                throw Error(record, "fuel khong dung dinh dang.");
            decimal quantity = ParseNonNegativeDecimal(record, parts[0], "fuel quantity");
            if (quantity <= 0m)
                throw Error(record, "fuel quantity phai lon hon 0.");
            string unit = parts[1];
            string code = string.Join("-", parts.Skip(2));
            MachineFuelKind kind;
            if (code.StartsWith("pin-", StringComparison.Ordinal))
                kind = MachineFuelKind.Battery;
            else if (code.StartsWith("diesel", StringComparison.Ordinal))
                kind = MachineFuelKind.Diesel;
            else if (code.StartsWith("gasoline", StringComparison.Ordinal))
                kind = MachineFuelKind.Gasoline;
            else if (code.StartsWith("electric", StringComparison.Ordinal))
                kind = MachineFuelKind.Electricity;
            else
                kind = MachineFuelKind.Other;
            return new MachineFuelRequirement(quantity, unit, code, kind);
        }

        private static IEnumerable<MachineOperatorRequirement> ParseOperators(
            RegulationDataRecord record,
            string value)
        {
            if (string.Equals(value, "none", StringComparison.Ordinal))
                return Array.Empty<MachineOperatorRequirement>();
            var result = new List<MachineOperatorRequirement>();
            foreach (string item in value.Split('+'))
            {
                int separator = item.IndexOf('-');
                if (separator <= 0 || separator == item.Length - 1)
                    throw Error(record, "operators khong dung dinh dang.");
                decimal quantity = ParseNonNegativeDecimal(
                    record, item.Substring(0, separator), "operator quantity");
                if (quantity <= 0m)
                    throw Error(record, "operator quantity phai lon hon 0.");
                result.Add(new MachineOperatorRequirement(
                    quantity,
                    item.Substring(separator + 1)));
            }
            return result;
        }

        private static string DeriveCode(string key, bool nonState)
        {
            const string prefix = "MACHINE-M010.";
            if (string.IsNullOrEmpty(key) || !key.StartsWith(prefix, StringComparison.Ordinal))
                throw new FormatException("MachineRate key khong dung MACHINE-M010.xxx.");
            return (nonState ? "M011." : "M010.") + key.Substring(prefix.Length);
        }

        private static string Required(
            IReadOnlyDictionary<string, string> fields,
            RegulationDataRecord record,
            string key)
        {
            if (!fields.TryGetValue(key, out string value) || string.IsNullOrWhiteSpace(value))
                throw Error(record, "thieu field " + key + ".");
            return value;
        }

        private static string Select(
            IReadOnlyDictionary<string, string> fields,
            string overrideKey,
            string defaultValue)
        {
            return !string.IsNullOrEmpty(overrideKey) &&
                fields.TryGetValue(overrideKey, out string value)
                ? value
                : defaultValue;
        }

        private static int ParseInt(RegulationDataRecord record, string value, string name)
        {
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int result))
                throw Error(record, name + " khong phai so nguyen invariant.");
            return result;
        }

        private static decimal ParseNonNegativeDecimal(
            RegulationDataRecord record,
            string value,
            string name)
        {
            if (!decimal.TryParse(
                value,
                NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture,
                out decimal result) ||
                result < 0m)
            {
                throw Error(record, name + " khong phai so invariant khong am.");
            }
            return result;
        }

        private static FormatException Error(RegulationDataRecord record, string message)
        {
            return new FormatException((record?.Key ?? "MachineRate") + ": " + message);
        }
    }
}
