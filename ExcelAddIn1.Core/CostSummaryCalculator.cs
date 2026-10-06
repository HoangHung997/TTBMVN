using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum CostSummaryTemplate
    {
        Survey = 1,
        IndependentStateFunded = 2,
        ProjectItemStateFunded = 3,
        OtherFunding = 4
    }

    public sealed class CostSummaryCalculationRequest
    {
        public CostSummaryCalculationRequest(
            CostSummaryTemplate template,
            decimal materialVnd,
            decimal laborVnd,
            decimal machineVnd,
            string terrain,
            decimal areaHa,
            CostProjectKind projectKind,
            CostConstructionKind constructionKind,
            decimal disposalWeightKg,
            decimal preTaxIncomeRatePercent,
            decimal vatRatePercent,
            CostComponentSelection selectedComponents,
            IEnumerable<CostRuleOverride> overrides = null)
        {
            Template = template;
            MaterialVnd = materialVnd;
            LaborVnd = laborVnd;
            MachineVnd = machineVnd;
            Terrain = (terrain ?? string.Empty).Trim();
            AreaHa = areaHa;
            ProjectKind = projectKind;
            ConstructionKind = constructionKind;
            DisposalWeightKg = disposalWeightKg;
            PreTaxIncomeRatePercent = preTaxIncomeRatePercent;
            VatRatePercent = vatRatePercent;
            SelectedComponents = selectedComponents;
            Overrides = new ReadOnlyCollection<CostRuleOverride>(
                (overrides ?? Enumerable.Empty<CostRuleOverride>())
                    .Where(item => item != null)
                    .ToList());
        }

        public CostSummaryTemplate Template { get; }
        public decimal MaterialVnd { get; }
        public decimal LaborVnd { get; }
        public decimal MachineVnd { get; }
        public string Terrain { get; }
        public decimal AreaHa { get; }
        public CostProjectKind ProjectKind { get; }
        public CostConstructionKind ConstructionKind { get; }
        public decimal DisposalWeightKg { get; }
        public decimal PreTaxIncomeRatePercent { get; }
        public decimal VatRatePercent { get; }
        public CostComponentSelection SelectedComponents { get; }
        public IReadOnlyList<CostRuleOverride> Overrides { get; }
    }

    public sealed class CostSummaryCalculationResult
    {
        internal CostSummaryCalculationResult(
            CostSummaryTemplate template,
            DirectCostResult directCost,
            long preTaxIncomeVnd,
            long zVnd,
            IEnumerable<CostComponentCalculationResult> components,
            long otherCostTotalVnd,
            long beforeTaxVnd,
            long taxableBaseVnd,
            long vatVnd,
            long afterTaxVnd,
            long roundedAfterTaxVnd)
        {
            Template = template;
            DirectCost = directCost;
            PreTaxIncomeVnd = preTaxIncomeVnd;
            ZVnd = zVnd;
            Components = new ReadOnlyCollection<CostComponentCalculationResult>(
                components.OrderBy(item => item.ComponentCode, StringComparer.Ordinal).ToList());
            OtherCostTotalVnd = otherCostTotalVnd;
            BeforeTaxVnd = beforeTaxVnd;
            TaxableBaseVnd = taxableBaseVnd;
            VatVnd = vatVnd;
            AfterTaxVnd = afterTaxVnd;
            RoundedAfterTaxVnd = roundedAfterTaxVnd;
        }

        public CostSummaryTemplate Template { get; }
        public DirectCostResult DirectCost { get; }
        public long PreTaxIncomeVnd { get; }
        public long ZVnd { get; }
        public IReadOnlyList<CostComponentCalculationResult> Components { get; }
        public long OtherCostTotalVnd { get; }
        public long BeforeTaxVnd { get; }
        public long TaxableBaseVnd { get; }
        public long VatVnd { get; }
        public long AfterTaxVnd { get; }
        public long RoundedAfterTaxVnd { get; }
        public string MoneyRoundingRule => "nearest-vnd-away-from-zero";
        public string FinalRoundingRule => "nearest-1000-vnd-away-from-zero";

        public CostComponentCalculationResult FindRequired(string componentCode)
        {
            string key = (componentCode ?? string.Empty).Trim().ToUpperInvariant();
            CostComponentCalculationResult result = Components.FirstOrDefault(item =>
                string.Equals(item.ComponentCode, key, StringComparison.Ordinal));
            if (result == null)
                throw new KeyNotFoundException("Khong co ket qua " + componentCode + ".");
            return result;
        }
    }

    public static class CostSummaryCalculator
    {
        private static readonly ISet<string> ComponentCodes = new HashSet<string>(
            new[] { "K1", "K2", "K3", "K4", "K5", "K6" },
            StringComparer.Ordinal);

        public static CostSummaryCalculationResult Calculate(
            CostRuleCatalog catalog,
            CostSummaryCalculationRequest request)
        {
            if (catalog == null)
                throw new ArgumentNullException(nameof(catalog));
            if (request == null)
                throw new ArgumentNullException(nameof(request));
            ValidateRequest(request);
            Dictionary<string, CostRuleOverride> overrides = ValidateOverrides(request);

            DirectCostResult direct = CostRuleCalculator.CalculateDirect(
                catalog,
                request.MaterialVnd,
                request.LaborVnd,
                request.MachineVnd);
            long preTaxIncome = request.Template == CostSummaryTemplate.OtherFunding
                ? CostRuleCalculator.RoundMoney(
                    (direct.DirectVnd + direct.CommonVnd) * request.PreTaxIncomeRatePercent / 100m)
                : 0L;
            long z = checked(direct.ZVnd + preTaxIncome);
            var components = new List<CostComponentCalculationResult>();

            TerrainCostResult terrain = null;
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K1) ||
                IsSelected(request.SelectedComponents, CostComponentSelection.K4))
            {
                terrain = CostRuleCalculator.CalculateK1K4(
                    catalog,
                    request.Terrain,
                    z,
                    request.AreaHa);
            }

            if (IsSelected(request.SelectedComponents, CostComponentSelection.K1))
                components.Add(Apply("K1", terrain.K1Percent, terrain.K1Vnd, terrain.Source, string.Empty, overrides));
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K2))
            {
                CostRateAmountResult value = CostRuleCalculator.CalculateK2(
                    catalog,
                    request.ProjectKind,
                    direct.DirectVnd);
                components.Add(Apply(
                    "K2", value.RatePercent, value.AmountVnd, value.Source,
                    catalog.GetK2(request.ProjectKind).CurrentExternalBasis, overrides));
            }
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K3))
            {
                CostRateAmountResult value = CostRuleCalculator.CalculateK3(catalog, z);
                components.Add(Apply("K3", value.RatePercent, value.AmountVnd, value.Source, string.Empty, overrides));
            }
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K4))
                components.Add(Apply("K4", terrain.K4Percent, terrain.K4Vnd, terrain.Source, string.Empty, overrides));
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K5))
            {
                CostRateAmountResult value = CostRuleCalculator.CalculateK5(
                    catalog,
                    request.ConstructionKind,
                    z);
                components.Add(Apply(
                    "K5", value.RatePercent, value.AmountVnd, value.Source,
                    catalog.GetK5(request.ConstructionKind).CurrentExternalBasis, overrides));
            }
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K6))
            {
                CostRateAmountResult value = CostRuleCalculator.CalculateK6(
                    catalog,
                    z,
                    request.DisposalWeightKg,
                    request.AreaHa);
                components.Add(Apply("K6", value.RatePercent, value.AmountVnd, value.Source, string.Empty, overrides));
            }

            long otherTotal = components.Aggregate(
                0L,
                (sum, item) => checked(sum + item.AppliedAmountVnd));
            long beforeTax = checked(z + otherTotal);
            long taxableBase = beforeTax;
            long vat = 0L;
            if (request.Template == CostSummaryTemplate.OtherFunding)
            {
                taxableBase = checked(
                    beforeTax - AppliedAmount(components, "K3") - AppliedAmount(components, "K4"));
                if (taxableBase < 0)
                    throw new InvalidOperationException("Co so tinh VAT am sau khi loai K3 va K4.");
                vat = CostRuleCalculator.RoundMoney(taxableBase * request.VatRatePercent / 100m);
            }
            long afterTax = checked(beforeTax + vat);
            long rounded = RoundToThousand(afterTax);

            return new CostSummaryCalculationResult(
                request.Template,
                direct,
                preTaxIncome,
                z,
                components,
                otherTotal,
                beforeTax,
                taxableBase,
                vat,
                afterTax,
                rounded);
        }

        public static long RoundToThousand(decimal value)
        {
            long thousands = CostRuleCalculator.RoundMoney(value / 1000m);
            return checked(thousands * 1000L);
        }

        private static CostComponentCalculationResult Apply(
            string code,
            decimal rate,
            long calculated,
            RegulationSourceLocator source,
            string externalBasis,
            IDictionary<string, CostRuleOverride> overrides)
        {
            if (!overrides.TryGetValue(code, out CostRuleOverride ruleOverride))
            {
                return new CostComponentCalculationResult(
                    code, rate, calculated, calculated, string.Empty, source, externalBasis);
            }
            return new CostComponentCalculationResult(
                code,
                rate,
                calculated,
                ruleOverride.AmountVnd,
                ruleOverride.Reason,
                source,
                externalBasis);
        }

        private static long AppliedAmount(
            IEnumerable<CostComponentCalculationResult> components,
            string code)
        {
            CostComponentCalculationResult item = components.FirstOrDefault(value =>
                string.Equals(value.ComponentCode, code, StringComparison.Ordinal));
            return item == null ? 0L : item.AppliedAmountVnd;
        }

        private static bool IsSelected(CostComponentSelection selection, CostComponentSelection value)
        {
            return (selection & value) == value;
        }

        private static void ValidateRequest(CostSummaryCalculationRequest request)
        {
            if (!Enum.IsDefined(typeof(CostSummaryTemplate), request.Template))
                throw new ArgumentOutOfRangeException(nameof(request.Template));
            if (request.SelectedComponents == CostComponentSelection.None ||
                (request.SelectedComponents & ~CostComponentSelection.All) != 0)
            {
                throw new ArgumentOutOfRangeException(nameof(request.SelectedComponents));
            }
            if (request.MaterialVnd < 0m || request.LaborVnd < 0m || request.MachineVnd < 0m ||
                request.AreaHa < 0m || request.DisposalWeightKg < 0m ||
                request.PreTaxIncomeRatePercent < 0m || request.VatRatePercent < 0m)
            {
                throw new ArgumentOutOfRangeException(nameof(request), "Gia tri tinh chi phi phai khong am.");
            }
        }

        private static Dictionary<string, CostRuleOverride> ValidateOverrides(
            CostSummaryCalculationRequest request)
        {
            var result = new Dictionary<string, CostRuleOverride>(StringComparer.Ordinal);
            foreach (CostRuleOverride item in request.Overrides)
            {
                if (!ComponentCodes.Contains(item.ComponentCode))
                    throw new ArgumentException("Ma ghi de khong hop le: " + item.ComponentCode + ".");
                CostComponentSelection flag = (CostComponentSelection)(1 << (item.ComponentCode[1] - '1'));
                if (!IsSelected(request.SelectedComponents, flag))
                    throw new ArgumentException("Khong the ghi de thanh phan chua chon: " + item.ComponentCode + ".");
                if (item.AmountVnd < 0 || item.Reason.Length == 0)
                    throw new ArgumentException("Ghi de " + item.ComponentCode + " phai co gia tri khong am va ly do.");
                if (result.ContainsKey(item.ComponentCode))
                    throw new ArgumentException("Trung ghi de " + item.ComponentCode + ".");
                result.Add(item.ComponentCode, item);
            }
            return result;
        }
    }
}
