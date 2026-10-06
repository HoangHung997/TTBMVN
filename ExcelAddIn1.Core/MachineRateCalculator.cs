using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ExcelAddIn1.Core
{
    public sealed class MachineRatePriceProfile
    {
        private readonly IReadOnlyDictionary<string, decimal> fuelUnitPricesVnd;
        private readonly IReadOnlyDictionary<string, decimal> laborDailyRatesVnd;
        private readonly IReadOnlyDictionary<MachineFuelKind, decimal> auxiliaryFactors;

        public MachineRatePriceProfile(
            IEnumerable<KeyValuePair<string, decimal>> fuelUnitPricesVnd,
            IEnumerable<KeyValuePair<string, decimal>> laborDailyRatesVnd,
            IEnumerable<KeyValuePair<MachineFuelKind, decimal>> auxiliaryFactorOverrides = null)
        {
            this.fuelUnitPricesVnd = CopyPrices(fuelUnitPricesVnd, "fuel");
            this.laborDailyRatesVnd = CopyPrices(laborDailyRatesVnd, "labor");
            var factors = new Dictionary<MachineFuelKind, decimal>
            {
                { MachineFuelKind.None, 1m },
                { MachineFuelKind.Battery, 1m },
                { MachineFuelKind.Diesel, 1.03m },
                { MachineFuelKind.Gasoline, 1.02m },
                { MachineFuelKind.Electricity, 1.05m },
                { MachineFuelKind.Other, 1m }
            };
            if (auxiliaryFactorOverrides != null)
            {
                foreach (KeyValuePair<MachineFuelKind, decimal> item in auxiliaryFactorOverrides)
                {
                    if (!Enum.IsDefined(typeof(MachineFuelKind), item.Key) || item.Value <= 0m)
                        throw new ArgumentException("He so nhien lieu phu khong hop le.", nameof(auxiliaryFactorOverrides));
                    factors[item.Key] = item.Value;
                }
            }
            auxiliaryFactors = new ReadOnlyDictionary<MachineFuelKind, decimal>(factors);
        }

        internal decimal GetFuelUnitPrice(string code)
        {
            if (!fuelUnitPricesVnd.TryGetValue(code, out decimal value))
                throw new KeyNotFoundException("Thieu gia nhien lieu/nang luong: " + code + ".");
            return value;
        }

        internal decimal GetLaborDailyRate(string code)
        {
            if (!laborDailyRatesVnd.TryGetValue(code, out decimal value))
                throw new KeyNotFoundException("Thieu don gia nhan cong: " + code + ".");
            return value;
        }

        internal decimal GetAuxiliaryFactor(MachineFuelKind kind)
        {
            return auxiliaryFactors[kind];
        }

        private static IReadOnlyDictionary<string, decimal> CopyPrices(
            IEnumerable<KeyValuePair<string, decimal>> values,
            string name)
        {
            var result = new Dictionary<string, decimal>(StringComparer.Ordinal);
            if (values != null)
            {
                foreach (KeyValuePair<string, decimal> item in values)
                {
                    string key = (item.Key ?? string.Empty).Trim();
                    if (string.IsNullOrEmpty(key) || item.Value < 0m || result.ContainsKey(key))
                        throw new ArgumentException("Bang gia " + name + " khong hop le.", name);
                    result.Add(key, item.Value);
                }
            }
            return new ReadOnlyDictionary<string, decimal>(result);
        }
    }

    public sealed class MachineRateCalculationOptions
    {
        public MachineRateCalculationOptions(
            bool corrosiveEnvironment = false,
            bool fuelCostIncludedInMaterials = false,
            decimal? originalPriceVnd = null,
            decimal? recoverableValuePercent = null)
        {
            if (originalPriceVnd.HasValue && originalPriceVnd.Value <= 0m)
                throw new ArgumentOutOfRangeException(nameof(originalPriceVnd));
            if (recoverableValuePercent.HasValue &&
                (recoverableValuePercent.Value < 0m || recoverableValuePercent.Value > 100m))
            {
                throw new ArgumentOutOfRangeException(nameof(recoverableValuePercent));
            }
            CorrosiveEnvironment = corrosiveEnvironment;
            FuelCostIncludedInMaterials = fuelCostIncludedInMaterials;
            OriginalPriceVnd = originalPriceVnd;
            RecoverableValuePercent = recoverableValuePercent;
        }

        public bool CorrosiveEnvironment { get; }
        public bool FuelCostIncludedInMaterials { get; }
        public decimal? OriginalPriceVnd { get; }
        public decimal? RecoverableValuePercent { get; }
    }

    public sealed class MachineRateResult
    {
        internal MachineRateResult(
            decimal depreciationRawVnd,
            decimal repairRawVnd,
            decimal fuelRawVnd,
            decimal operatorLaborRawVnd,
            decimal otherRawVnd)
        {
            DepreciationRawVnd = depreciationRawVnd;
            RepairRawVnd = repairRawVnd;
            FuelRawVnd = fuelRawVnd;
            OperatorLaborRawVnd = operatorLaborRawVnd;
            OtherRawVnd = otherRawVnd;
            RawTotalVnd = depreciationRawVnd + repairRawVnd + fuelRawVnd +
                operatorLaborRawVnd + otherRawVnd;

            DepreciationVnd = MachineRateCalculator.RoundVnd(depreciationRawVnd);
            RepairVnd = MachineRateCalculator.RoundVnd(repairRawVnd);
            FuelVnd = MachineRateCalculator.RoundVnd(fuelRawVnd);
            OperatorLaborVnd = MachineRateCalculator.RoundVnd(operatorLaborRawVnd);
            OtherVnd = MachineRateCalculator.RoundVnd(otherRawVnd);
            TotalVnd = checked(
                DepreciationVnd + RepairVnd + FuelVnd + OperatorLaborVnd + OtherVnd);
            WaitingRawVnd = (depreciationRawVnd * 0.5m) +
                (operatorLaborRawVnd * 0.5m) + otherRawVnd;
            WaitingVnd = MachineRateCalculator.RoundVnd(WaitingRawVnd);
        }

        public decimal DepreciationRawVnd { get; }
        public decimal RepairRawVnd { get; }
        public decimal FuelRawVnd { get; }
        public decimal OperatorLaborRawVnd { get; }
        public decimal OtherRawVnd { get; }
        public decimal RawTotalVnd { get; }
        public long DepreciationVnd { get; }
        public long RepairVnd { get; }
        public long FuelVnd { get; }
        public long OperatorLaborVnd { get; }
        public long OtherVnd { get; }
        public long TotalVnd { get; }
        public decimal WaitingRawVnd { get; }
        public long WaitingVnd { get; }

        public long GetHourlyRateVnd(decimal workingHoursPerShift)
        {
            if (workingHoursPerShift <= 0m)
                throw new ArgumentOutOfRangeException(nameof(workingHoursPerShift));
            return MachineRateCalculator.RoundVnd(RawTotalVnd / workingHoursPerShift);
        }
    }

    public static class MachineRateCalculator
    {
        public const decimal RecoveryThresholdVnd = 30000000m;
        public const decimal StandardRecoveryPercent = 10m;
        public const decimal CorrosiveEnvironmentFactor = 1.05m;

        public static MachineRateResult Calculate(
            MachineRateDefinition definition,
            MachineRatePriceProfile priceProfile,
            MachineRateCalculationOptions options = null)
        {
            if (definition == null)
                throw new ArgumentNullException(nameof(definition));
            if (priceProfile == null)
                throw new ArgumentNullException(nameof(priceProfile));
            options = options ?? new MachineRateCalculationOptions();

            decimal originalPrice = options.OriginalPriceVnd ?? definition.ReferencePriceVnd;
            decimal recoveryPercent = options.RecoverableValuePercent ??
                (options.OriginalPriceVnd.HasValue
                    ? (originalPrice >= RecoveryThresholdVnd ? StandardRecoveryPercent : 0m)
                    : definition.ReferenceRecoverableValuePercent);
            decimal corrosionFactor = options.CorrosiveEnvironment
                ? CorrosiveEnvironmentFactor
                : 1m;
            decimal annualShifts = definition.AnnualShifts;
            decimal depreciation = (originalPrice - (originalPrice * recoveryPercent / 100m)) *
                (definition.DepreciationPercent * corrosionFactor / 100m) /
                annualShifts;
            decimal repair = originalPrice *
                (definition.RepairPercent * corrosionFactor / 100m) /
                annualShifts;
            decimal other = originalPrice *
                (definition.OtherPercent / 100m) /
                annualShifts;

            decimal fuel = 0m;
            if (!definition.Fuel.IsNone && !options.FuelCostIncludedInMaterials)
            {
                fuel = definition.Fuel.Quantity *
                    priceProfile.GetFuelUnitPrice(definition.Fuel.PriceCode) *
                    priceProfile.GetAuxiliaryFactor(definition.Fuel.Kind);
            }

            decimal labor = 0m;
            foreach (MachineOperatorRequirement requirement in definition.Operators)
            {
                labor += requirement.Quantity *
                    priceProfile.GetLaborDailyRate(requirement.LaborCode);
            }

            return new MachineRateResult(depreciation, repair, fuel, labor, other);
        }

        internal static long RoundVnd(decimal value)
        {
            return checked((long)decimal.Round(value, 0, MidpointRounding.AwayFromZero));
        }
    }
}
