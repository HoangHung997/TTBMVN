using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public sealed class DirectCostResult
    {
        internal DirectCostResult(long materialVnd, long laborVnd, long machineVnd, long commonVnd)
        {
            MaterialVnd = materialVnd;
            LaborVnd = laborVnd;
            MachineVnd = machineVnd;
            CommonVnd = commonVnd;
        }

        public long MaterialVnd { get; }
        public long LaborVnd { get; }
        public long MachineVnd { get; }
        public long DirectVnd => checked(MaterialVnd + LaborVnd + MachineVnd);
        public long CommonVnd { get; }
        public long ZVnd => checked(DirectVnd + CommonVnd);
    }

    public sealed class CostRateAmountResult
    {
        internal CostRateAmountResult(decimal ratePercent, long amountVnd, RegulationSourceLocator source)
        {
            RatePercent = ratePercent;
            AmountVnd = amountVnd;
            Source = source;
        }

        public decimal RatePercent { get; }
        public long AmountVnd { get; }
        public RegulationSourceLocator Source { get; }
        public string RoundingRule => "nearest-vnd-away-from-zero";
    }

    public sealed class TerrainCostResult
    {
        internal TerrainCostResult(
            decimal k1Percent,
            decimal k4Percent,
            long k1Vnd,
            long k4Vnd,
            RegulationSourceLocator source)
        {
            K1Percent = k1Percent;
            K4Percent = k4Percent;
            K1Vnd = k1Vnd;
            K4Vnd = k4Vnd;
            Source = source;
        }

        public decimal K1Percent { get; }
        public decimal K4Percent { get; }
        public long K1Vnd { get; }
        public long K4Vnd { get; }
        public long TotalVnd => checked(K1Vnd + K4Vnd);
        public RegulationSourceLocator Source { get; }
    }

    public sealed class EstimateSummaryResult
    {
        internal EstimateSummaryResult(
            long zVnd,
            IDictionary<string, long> otherCosts,
            long otherCostTotalVnd,
            long beforeTaxVnd,
            long taxableBaseVnd,
            decimal vatRatePercent,
            long vatVnd)
        {
            ZVnd = zVnd;
            OtherCosts = new ReadOnlyDictionary<string, long>(
                new Dictionary<string, long>(otherCosts, StringComparer.Ordinal));
            OtherCostTotalVnd = otherCostTotalVnd;
            BeforeTaxVnd = beforeTaxVnd;
            TaxableBaseVnd = taxableBaseVnd;
            VatRatePercent = vatRatePercent;
            VatVnd = vatVnd;
        }

        public long ZVnd { get; }
        public IReadOnlyDictionary<string, long> OtherCosts { get; }
        public long OtherCostTotalVnd { get; }
        public long BeforeTaxVnd { get; }
        public long TaxableBaseVnd { get; }
        public decimal VatRatePercent { get; }
        public long VatVnd { get; }
        public long AfterTaxVnd => checked(BeforeTaxVnd + VatVnd);
    }

    public static class CostRuleCalculator
    {
        private const decimal Billion = 1000000000m;

        public static DirectCostResult CalculateDirect(
            CostRuleCatalog catalog,
            decimal materialVnd,
            decimal laborVnd,
            decimal machineVnd)
        {
            Required(catalog);
            ValidateNonNegative(materialVnd, nameof(materialVnd));
            ValidateNonNegative(laborVnd, nameof(laborVnd));
            ValidateNonNegative(machineVnd, nameof(machineVnd));
            long material = RoundMoney(materialVnd);
            long labor = RoundMoney(laborVnd);
            long machine = RoundMoney(machineVnd);
            long common = RoundMoney(labor * catalog.CommonRatePercent / 100m);
            return new DirectCostResult(material, labor, machine, common);
        }

        public static TerrainCostResult CalculateK1K4(
            CostRuleCatalog catalog,
            string terrain,
            decimal zVnd,
            decimal areaHa)
        {
            Required(catalog);
            ValidateNonNegative(zVnd, nameof(zVnd));
            ValidateNonNegative(areaHa, nameof(areaHa));
            CostTerrainDefinition definition = catalog.FindTerrain(terrain);
            long k1 = RoundMoney(zVnd * definition.K1Percent / 100m);
            if (areaHa <= catalog.MinimumAreaHa)
                k1 = Math.Max(k1, catalog.MinimumSurveyVnd);
            long k4 = RoundMoney(zVnd * definition.K4Percent / 100m);
            return new TerrainCostResult(
                definition.K1Percent,
                definition.K4Percent,
                k1,
                k4,
                definition.Source);
        }

        public static CostRateAmountResult CalculateK2(
            CostRuleCatalog catalog,
            CostProjectKind projectKind,
            decimal directVnd)
        {
            Required(catalog);
            ValidateNonNegative(directVnd, nameof(directVnd));
            CostRateTableDefinition table = catalog.GetK2(projectKind);
            decimal billion = directVnd / Billion;
            decimal rate = table.RatesPercent[table.RatesPercent.Count - 1];
            for (int index = 0; index < table.ThresholdsBillion.Count; index++)
            {
                if (billion <= table.ThresholdsBillion[index])
                {
                    rate = table.RatesPercent[index];
                    break;
                }
            }
            return RateAmount(rate, directVnd, table.Source);
        }

        public static CostRateAmountResult CalculateK3(CostRuleCatalog catalog, decimal zVnd)
        {
            Required(catalog);
            ValidateNonNegative(zVnd, nameof(zVnd));
            decimal billion = zVnd / Billion;
            CostBracketDefinition bracket = catalog.K3Brackets.SingleOrDefault(item =>
                billion >= item.MinimumBillion &&
                (!item.MaximumBillionExclusive.HasValue || billion < item.MaximumBillionExclusive.Value));
            if (bracket == null)
                throw new InvalidOperationException("Khong tim thay khoang K3.");
            long amount = RoundMoney(zVnd * bracket.RatePercent / 100m);
            amount = Math.Max(catalog.K3MinimumVnd, Math.Min(catalog.K3MaximumVnd, amount));
            return new CostRateAmountResult(bracket.RatePercent, amount, bracket.Source);
        }

        public static CostRateAmountResult CalculateK5(
            CostRuleCatalog catalog,
            CostConstructionKind constructionKind,
            decimal zVnd)
        {
            Required(catalog);
            ValidateNonNegative(zVnd, nameof(zVnd));
            CostRateTableDefinition table = catalog.GetK5(constructionKind);
            decimal billion = zVnd / Billion;
            decimal rate;
            if (billion <= table.ThresholdsBillion[0])
            {
                rate = table.RatesPercent[0];
            }
            else
            {
                int upper = -1;
                for (int index = 1; index < table.ThresholdsBillion.Count; index++)
                {
                    if (billion <= table.ThresholdsBillion[index])
                    {
                        upper = index;
                        break;
                    }
                }
                if (upper < 0)
                    throw new InvalidOperationException("K5 vuot moc toi da; phai lap du toan chi phi rieng.");
                decimal lowerThreshold = table.ThresholdsBillion[upper - 1];
                decimal upperThreshold = table.ThresholdsBillion[upper];
                decimal lowerRate = table.RatesPercent[upper - 1];
                decimal upperRate = table.RatesPercent[upper];
                rate = lowerRate -
                    ((lowerRate - upperRate) / (upperThreshold - lowerThreshold)) *
                    (billion - lowerThreshold);
            }
            return RateAmount(rate, zVnd, table.Source);
        }

        public static CostRateAmountResult CalculateK6(
            CostRuleCatalog catalog,
            decimal zVnd,
            decimal disposalWeightKg,
            decimal areaHa)
        {
            Required(catalog);
            ValidateNonNegative(zVnd, nameof(zVnd));
            ValidateNonNegative(disposalWeightKg, nameof(disposalWeightKg));
            ValidateNonNegative(areaHa, nameof(areaHa));
            if (disposalWeightKg == 1000m)
                throw new InvalidOperationException("TT123 khong quy dinh muc K6 tai dung 1000 kg; can phe duyet cach ap dung.");
            decimal rate = disposalWeightKg < 1000m
                ? catalog.K6Under1000Percent
                : catalog.K6Over1000Percent;
            long amount = RoundMoney(zVnd * rate / 100m);
            if (areaHa <= catalog.MinimumAreaHa)
                amount = Math.Max(amount, catalog.MinimumDisposalVnd);
            return new CostRateAmountResult(rate, amount, catalog.K6Source);
        }

        public static EstimateSummaryResult CalculateSummary(
            decimal zVnd,
            IEnumerable<KeyValuePair<string, long>> otherCosts,
            decimal vatRatePercent)
        {
            ValidateNonNegative(zVnd, nameof(zVnd));
            ValidateNonNegative(vatRatePercent, nameof(vatRatePercent));
            var values = new Dictionary<string, long>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, long> item in otherCosts ?? Enumerable.Empty<KeyValuePair<string, long>>())
            {
                string key = (item.Key ?? string.Empty).Trim().ToUpperInvariant();
                if (key.Length == 0 || item.Value < 0 || values.ContainsKey(key))
                    throw new ArgumentException("Chi phi K khong hop le: " + item.Key + ".", nameof(otherCosts));
                values.Add(key, item.Value);
            }
            long z = RoundMoney(zVnd);
            long otherTotal = values.Values.Aggregate(0L, checked((sum, value) => sum + value));
            long beforeTax = checked(z + otherTotal);
            values.TryGetValue("K3", out long k3);
            values.TryGetValue("K4", out long k4);
            long taxableBase = checked(beforeTax - k3 - k4);
            if (taxableBase < 0)
                throw new InvalidOperationException("Co so tinh VAT am sau khi loai K3 va K4.");
            long vat = RoundMoney(taxableBase * vatRatePercent / 100m);
            return new EstimateSummaryResult(
                z,
                values,
                otherTotal,
                beforeTax,
                taxableBase,
                vatRatePercent,
                vat);
        }

        public static long RoundMoney(decimal value)
        {
            decimal rounded = decimal.Round(value, 0, MidpointRounding.AwayFromZero);
            if (rounded > long.MaxValue || rounded < long.MinValue)
                throw new OverflowException("Gia tri tien vuot gioi han Int64.");
            return decimal.ToInt64(rounded);
        }

        private static CostRateAmountResult RateAmount(
            decimal rate,
            decimal basisVnd,
            RegulationSourceLocator source)
        {
            return new CostRateAmountResult(rate, RoundMoney(basisVnd * rate / 100m), source);
        }

        private static void Required(CostRuleCatalog catalog)
        {
            if (catalog == null)
                throw new ArgumentNullException(nameof(catalog));
        }

        private static void ValidateNonNegative(decimal value, string parameter)
        {
            if (value < 0m)
                throw new ArgumentOutOfRangeException(parameter, "Gia tri phai khong am.");
        }
    }
}
