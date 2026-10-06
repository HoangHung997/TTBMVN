using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    [Flags]
    public enum CostComponentSelection
    {
        None = 0,
        K1 = 1,
        K2 = 2,
        K3 = 4,
        K4 = 8,
        K5 = 16,
        K6 = 32,
        All = K1 | K2 | K3 | K4 | K5 | K6
    }

    public sealed class CostRuleOverride
    {
        public CostRuleOverride(string componentCode, long amountVnd, string reason)
        {
            ComponentCode = (componentCode ?? string.Empty).Trim().ToUpperInvariant();
            AmountVnd = amountVnd;
            Reason = (reason ?? string.Empty).Trim();
        }

        public string ComponentCode { get; }
        public long AmountVnd { get; }
        public string Reason { get; }
    }

    public sealed class CostRuleCalculationRequest
    {
        public CostRuleCalculationRequest(
            decimal materialVnd,
            decimal laborVnd,
            decimal machineVnd,
            string terrain,
            decimal areaHa,
            CostProjectKind projectKind,
            CostConstructionKind constructionKind,
            decimal disposalWeightKg,
            decimal vatRatePercent,
            CostComponentSelection selectedComponents,
            IEnumerable<CostRuleOverride> overrides = null)
        {
            MaterialVnd = materialVnd;
            LaborVnd = laborVnd;
            MachineVnd = machineVnd;
            Terrain = (terrain ?? string.Empty).Trim();
            AreaHa = areaHa;
            ProjectKind = projectKind;
            ConstructionKind = constructionKind;
            DisposalWeightKg = disposalWeightKg;
            VatRatePercent = vatRatePercent;
            SelectedComponents = selectedComponents;
            Overrides = new ReadOnlyCollection<CostRuleOverride>(
                (overrides ?? Enumerable.Empty<CostRuleOverride>())
                    .Where(item => item != null)
                    .ToList());
        }

        public decimal MaterialVnd { get; }
        public decimal LaborVnd { get; }
        public decimal MachineVnd { get; }
        public string Terrain { get; }
        public decimal AreaHa { get; }
        public CostProjectKind ProjectKind { get; }
        public CostConstructionKind ConstructionKind { get; }
        public decimal DisposalWeightKg { get; }
        public decimal VatRatePercent { get; }
        public CostComponentSelection SelectedComponents { get; }
        public IReadOnlyList<CostRuleOverride> Overrides { get; }
    }

    public sealed class CostComponentCalculationResult
    {
        internal CostComponentCalculationResult(
            string componentCode,
            decimal? calculatedRatePercent,
            long calculatedAmountVnd,
            long appliedAmountVnd,
            string overrideReason,
            RegulationSourceLocator source,
            string externalBasis)
        {
            ComponentCode = componentCode;
            CalculatedRatePercent = calculatedRatePercent;
            CalculatedAmountVnd = calculatedAmountVnd;
            AppliedAmountVnd = appliedAmountVnd;
            OverrideReason = overrideReason ?? string.Empty;
            Source = source;
            ExternalBasis = externalBasis ?? string.Empty;
        }

        public string ComponentCode { get; }
        public decimal? CalculatedRatePercent { get; }
        public long CalculatedAmountVnd { get; }
        public long AppliedAmountVnd { get; }
        public bool IsOverridden => OverrideReason.Length > 0;
        public string OverrideReason { get; }
        public RegulationSourceLocator Source { get; }
        public string ExternalBasis { get; }
        public string RoundingRule => "nearest-vnd-away-from-zero";
    }

    public sealed class CostRuleEngineResult
    {
        internal CostRuleEngineResult(
            DirectCostResult directCost,
            IEnumerable<CostComponentCalculationResult> components,
            EstimateSummaryResult summary)
        {
            DirectCost = directCost;
            Components = new ReadOnlyCollection<CostComponentCalculationResult>(
                components.OrderBy(item => item.ComponentCode, StringComparer.Ordinal).ToList());
            Summary = summary;
        }

        public DirectCostResult DirectCost { get; }
        public IReadOnlyList<CostComponentCalculationResult> Components { get; }
        public EstimateSummaryResult Summary { get; }

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

    public static class CostRuleEngine
    {
        private static readonly ISet<string> ComponentCodes = new HashSet<string>(
            new[] { "K1", "K2", "K3", "K4", "K5", "K6" },
            StringComparer.Ordinal);

        public static CostRuleEngineResult Calculate(
            CostRuleCatalog catalog,
            CostRuleCalculationRequest request)
        {
            if (catalog == null)
                throw new ArgumentNullException(nameof(catalog));
            if (request == null)
                throw new ArgumentNullException(nameof(request));
            ValidateSelection(request.SelectedComponents);
            Dictionary<string, CostRuleOverride> overrides = ValidateOverrides(
                request.Overrides,
                request.SelectedComponents);

            DirectCostResult direct = CostRuleCalculator.CalculateDirect(
                catalog,
                request.MaterialVnd,
                request.LaborVnd,
                request.MachineVnd);
            var components = new List<CostComponentCalculationResult>();

            if (IsSelected(request.SelectedComponents, CostComponentSelection.K1) ||
                IsSelected(request.SelectedComponents, CostComponentSelection.K4))
            {
                TerrainCostResult terrain = CostRuleCalculator.CalculateK1K4(
                    catalog,
                    request.Terrain,
                    direct.ZVnd,
                    request.AreaHa);
                if (IsSelected(request.SelectedComponents, CostComponentSelection.K1))
                    components.Add(Apply("K1", terrain.K1Percent, terrain.K1Vnd, terrain.Source, string.Empty, overrides));
                if (IsSelected(request.SelectedComponents, CostComponentSelection.K4))
                    components.Add(Apply("K4", terrain.K4Percent, terrain.K4Vnd, terrain.Source, string.Empty, overrides));
            }
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K2))
            {
                CostRateAmountResult value = CostRuleCalculator.CalculateK2(
                    catalog,
                    request.ProjectKind,
                    direct.DirectVnd);
                components.Add(Apply(
                    "K2",
                    value.RatePercent,
                    value.AmountVnd,
                    value.Source,
                    catalog.GetK2(request.ProjectKind).CurrentExternalBasis,
                    overrides));
            }
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K3))
            {
                CostRateAmountResult value = CostRuleCalculator.CalculateK3(catalog, direct.ZVnd);
                components.Add(Apply("K3", value.RatePercent, value.AmountVnd, value.Source, string.Empty, overrides));
            }
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K5))
            {
                CostRateAmountResult value = CostRuleCalculator.CalculateK5(
                    catalog,
                    request.ConstructionKind,
                    direct.ZVnd);
                components.Add(Apply(
                    "K5",
                    value.RatePercent,
                    value.AmountVnd,
                    value.Source,
                    catalog.GetK5(request.ConstructionKind).CurrentExternalBasis,
                    overrides));
            }
            if (IsSelected(request.SelectedComponents, CostComponentSelection.K6))
            {
                CostRateAmountResult value = CostRuleCalculator.CalculateK6(
                    catalog,
                    direct.ZVnd,
                    request.DisposalWeightKg,
                    request.AreaHa);
                components.Add(Apply("K6", value.RatePercent, value.AmountVnd, value.Source, string.Empty, overrides));
            }

            EstimateSummaryResult summary = CostRuleCalculator.CalculateSummary(
                direct.ZVnd,
                components.Select(component => new KeyValuePair<string, long>(
                    component.ComponentCode,
                    component.AppliedAmountVnd)),
                request.VatRatePercent);
            return new CostRuleEngineResult(direct, components, summary);
        }

        private static CostComponentCalculationResult Apply(
            string code,
            decimal rate,
            long calculatedAmount,
            RegulationSourceLocator source,
            string externalBasis,
            IDictionary<string, CostRuleOverride> overrides)
        {
            CostRuleOverride value;
            if (overrides.TryGetValue(code, out value))
            {
                return new CostComponentCalculationResult(
                    code,
                    rate,
                    calculatedAmount,
                    value.AmountVnd,
                    value.Reason,
                    source,
                    externalBasis);
            }
            return new CostComponentCalculationResult(
                code,
                rate,
                calculatedAmount,
                calculatedAmount,
                string.Empty,
                source,
                externalBasis);
        }

        private static Dictionary<string, CostRuleOverride> ValidateOverrides(
            IEnumerable<CostRuleOverride> values,
            CostComponentSelection selection)
        {
            var result = new Dictionary<string, CostRuleOverride>(StringComparer.Ordinal);
            foreach (CostRuleOverride value in values ?? Enumerable.Empty<CostRuleOverride>())
            {
                if (value == null)
                    continue;
                if (!ComponentCodes.Contains(value.ComponentCode))
                    throw new ArgumentException("Ma override khong hop le: " + value.ComponentCode + ".");
                CostComponentSelection flag = (CostComponentSelection)Enum.Parse(
                    typeof(CostComponentSelection),
                    value.ComponentCode,
                    false);
                if (!IsSelected(selection, flag))
                    throw new ArgumentException("Override cho thanh phan khong duoc chon: " + value.ComponentCode + ".");
                if (value.AmountVnd < 0)
                    throw new ArgumentOutOfRangeException(nameof(values), "Gia tri override khong duoc am.");
                if (value.Reason.Length == 0 || value.Reason.Length > 1024)
                    throw new ArgumentException("Override phai co ly do tu 1 den 1024 ky tu.", nameof(values));
                if (result.ContainsKey(value.ComponentCode))
                    throw new ArgumentException("Override bi trung: " + value.ComponentCode + ".", nameof(values));
                result.Add(value.ComponentCode, value);
            }
            return result;
        }

        private static void ValidateSelection(CostComponentSelection selection)
        {
            if (selection == CostComponentSelection.None ||
                (selection & ~CostComponentSelection.All) != 0)
            {
                throw new ArgumentOutOfRangeException(nameof(selection), "Thanh phan chi phi khong hop le.");
            }
        }

        private static bool IsSelected(
            CostComponentSelection selection,
            CostComponentSelection value)
        {
            return (selection & value) == value;
        }
    }
}
