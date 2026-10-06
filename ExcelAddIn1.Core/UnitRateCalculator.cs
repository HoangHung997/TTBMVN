using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum UnitRateIssueCode
    {
        NormCalculation = 1,
        MissingPrice = 2,
        PriceKindMismatch = 3,
        UnitMismatch = 4,
        BindingRequired = 5,
        InvalidBinding = 6
    }

    public sealed class UnitRateResourceBinding
    {
        public UnitRateResourceBinding(string resourceCode, string priceCode, string reason)
        {
            ResourceCode = (resourceCode ?? string.Empty).Trim();
            PriceCode = (priceCode ?? string.Empty).Trim();
            Reason = (reason ?? string.Empty).Trim();
        }

        public string ResourceCode { get; }
        public string PriceCode { get; }
        public string Reason { get; }
    }

    public sealed class UnitRateCalculationRequest
    {
        public UnitRateCalculationRequest(
            NormDefinition definition,
            string variantCode,
            PriceProfile priceProfile,
            IEnumerable<string> conditions = null,
            IEnumerable<UnitRateResourceBinding> bindings = null)
        {
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
            VariantCode = (variantCode ?? string.Empty).Trim();
            PriceProfile = priceProfile ?? throw new ArgumentNullException(nameof(priceProfile));
            Conditions = new ReadOnlyCollection<string>(
                (conditions ?? Enumerable.Empty<string>())
                    .Select(value => (value ?? string.Empty).Trim())
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToList());
            Bindings = new ReadOnlyCollection<UnitRateResourceBinding>(
                (bindings ?? Enumerable.Empty<UnitRateResourceBinding>())
                    .Where(value => value != null)
                    .ToList());
        }

        public NormDefinition Definition { get; }
        public string VariantCode { get; }
        public PriceProfile PriceProfile { get; }
        public IReadOnlyList<string> Conditions { get; }
        public IReadOnlyList<UnitRateResourceBinding> Bindings { get; }
    }

    public sealed class UnitRateValidationIssue
    {
        internal UnitRateValidationIssue(UnitRateIssueCode issueCode, string resourceCode, string message)
        {
            IssueCode = issueCode;
            ResourceCode = resourceCode ?? string.Empty;
            Message = message ?? string.Empty;
        }

        public UnitRateIssueCode IssueCode { get; }
        public string ResourceCode { get; }
        public string Message { get; }
    }

    public sealed class UnitRateValidationResult
    {
        internal UnitRateValidationResult(
            NormCalculationResult normResult,
            IEnumerable<UnitRateValidationIssue> issues)
        {
            NormResult = normResult;
            Issues = new ReadOnlyCollection<UnitRateValidationIssue>(issues.ToList());
        }

        public NormCalculationResult NormResult { get; }
        public IReadOnlyList<UnitRateValidationIssue> Issues { get; }
        public bool IsValid => Issues.Count == 0;
    }

    public sealed class UnitRateValidationException : InvalidOperationException
    {
        internal UnitRateValidationException(UnitRateValidationResult validation)
            : base(string.Join(" ", validation.Issues.Select(issue => issue.Message)))
        {
            Validation = validation;
        }

        public UnitRateValidationResult Validation { get; }
    }

    public sealed class UnitRateResourceAmount
    {
        internal UnitRateResourceAmount(
            NormResourceKind kind,
            string resourceCode,
            string priceCode,
            string displayName,
            string unit,
            decimal quantity,
            decimal unitPriceVnd,
            decimal amountVnd,
            bool isPercentage,
            bool isPriceOverridden,
            string priceSourceReference,
            string bindingReason)
        {
            Kind = kind;
            ResourceCode = resourceCode;
            PriceCode = priceCode;
            DisplayName = displayName;
            Unit = unit;
            Quantity = quantity;
            UnitPriceVnd = unitPriceVnd;
            AmountVnd = amountVnd;
            IsPercentage = isPercentage;
            IsPriceOverridden = isPriceOverridden;
            PriceSourceReference = priceSourceReference;
            BindingReason = bindingReason;
        }

        public NormResourceKind Kind { get; }
        public string ResourceCode { get; }
        public string PriceCode { get; }
        public string DisplayName { get; }
        public string Unit { get; }
        public decimal Quantity { get; }
        public decimal UnitPriceVnd { get; }
        public decimal AmountVnd { get; }
        public bool IsPercentage { get; }
        public bool IsPriceOverridden { get; }
        public string PriceSourceReference { get; }
        public string BindingReason { get; }
    }

    public sealed class UnitRateCalculationResult
    {
        internal UnitRateCalculationResult(
            UnitRateCalculationRequest request,
            IEnumerable<UnitRateResourceAmount> resources,
            decimal materialAmountVnd,
            decimal laborAmountVnd,
            decimal machineAmountVnd)
        {
            NormKey = request.Definition.Key;
            NormTitle = request.Definition.Title;
            VariantCode = request.VariantCode;
            WorkUnit = request.Definition.WorkUnit;
            Conditions = request.Conditions;
            Bindings = request.Bindings;
            Resources = new ReadOnlyCollection<UnitRateResourceAmount>(resources.ToList());
            MaterialAmountVnd = materialAmountVnd;
            LaborAmountVnd = laborAmountVnd;
            MachineAmountVnd = machineAmountVnd;
            PriceProfileId = request.PriceProfile.ProfileId;
            PriceProfileVersion = request.PriceProfile.DataVersion;
            PriceProfileChecksum = request.PriceProfile.Checksum;
            NormSource = request.Definition.Source;
        }

        public string NormKey { get; }
        public string NormTitle { get; }
        public string VariantCode { get; }
        public string WorkUnit { get; }
        public IReadOnlyList<string> Conditions { get; }
        public IReadOnlyList<UnitRateResourceBinding> Bindings { get; }
        public IReadOnlyList<UnitRateResourceAmount> Resources { get; }
        public decimal MaterialAmountVnd { get; }
        public decimal LaborAmountVnd { get; }
        public decimal MachineAmountVnd { get; }
        public decimal TotalAmountVnd => MaterialAmountVnd + LaborAmountVnd + MachineAmountVnd;
        public string PriceProfileId { get; }
        public string PriceProfileVersion { get; }
        public string PriceProfileChecksum { get; }
        public RegulationSourceLocator NormSource { get; }
        public string RoundingRule => "none-invariant-decimal";
    }

    public static class UnitRateCalculator
    {
        public static UnitRateValidationResult Validate(UnitRateCalculationRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            var issues = new List<UnitRateValidationIssue>();
            NormCalculationResult normResult = null;
            try
            {
                normResult = NormCalculator.Calculate(
                    request.Definition,
                    request.VariantCode,
                    1m,
                    request.Conditions);
            }
            catch (Exception ex) when (
                ex is ArgumentException ||
                ex is InvalidOperationException ||
                ex is KeyNotFoundException)
            {
                issues.Add(new UnitRateValidationIssue(
                    UnitRateIssueCode.NormCalculation,
                    request.Definition.Key,
                    ex.Message));
                return new UnitRateValidationResult(null, issues);
            }

            Dictionary<string, UnitRateResourceBinding> bindings = ValidateBindings(
                request,
                normResult,
                issues);
            foreach (NormConsumption consumption in normResult.Consumptions)
            {
                if (consumption.Quantity == 0m)
                    continue;
                if (IsPercentage(consumption))
                    continue;

                UnitRateResourceBinding binding;
                bindings.TryGetValue(consumption.ResourceCode, out binding);
                PriceProfilePrice price;
                if (request.PriceProfile.TryFind(consumption.ResourceCode, out price))
                {
                    if (binding != null)
                    {
                        issues.Add(new UnitRateValidationIssue(
                            UnitRateIssueCode.InvalidBinding,
                            consumption.ResourceCode,
                            "Khong duoc gan binding cho ma da co gia truc tiep: " +
                                consumption.ResourceCode + "."));
                    }
                }
                else if (binding == null)
                {
                    bool logical = IsLogicalResource(consumption.ResourceCode);
                    issues.Add(new UnitRateValidationIssue(
                        logical ? UnitRateIssueCode.BindingRequired : UnitRateIssueCode.MissingPrice,
                        consumption.ResourceCode,
                        logical
                            ? "Phai chon ma gia cu the cho nguon luc logic " + consumption.ResourceCode + "."
                            : "Thieu gia cho ma " + consumption.ResourceCode + "."));
                    continue;
                }
                else if (!request.PriceProfile.TryFind(binding.PriceCode, out price))
                {
                    issues.Add(new UnitRateValidationIssue(
                        UnitRateIssueCode.MissingPrice,
                        consumption.ResourceCode,
                        "Binding " + consumption.ResourceCode + " tro den ma gia khong ton tai: " +
                            binding.PriceCode + "."));
                    continue;
                }

                if (price == null)
                    continue;
                PriceResourceKind expectedKind = ToPriceKind(consumption.Kind);
                if (price.Entry.Kind != expectedKind)
                {
                    issues.Add(new UnitRateValidationIssue(
                        UnitRateIssueCode.PriceKindMismatch,
                        consumption.ResourceCode,
                        "Loai gia khong khop " + consumption.ResourceCode + ": can " +
                            expectedKind + ", co " + price.Entry.Kind + "."));
                }
                if (!string.Equals(price.Entry.Unit, consumption.Unit, StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add(new UnitRateValidationIssue(
                        UnitRateIssueCode.UnitMismatch,
                        consumption.ResourceCode,
                        "Don vi gia khong khop " + consumption.ResourceCode + ": can " +
                            consumption.Unit + ", co " + price.Entry.Unit + "."));
                }
            }
            return new UnitRateValidationResult(normResult, issues);
        }

        public static UnitRateCalculationResult Calculate(UnitRateCalculationRequest request)
        {
            UnitRateValidationResult validation = Validate(request);
            if (!validation.IsValid)
                throw new UnitRateValidationException(validation);

            var bindings = request.Bindings.ToDictionary(
                item => item.ResourceCode,
                StringComparer.OrdinalIgnoreCase);
            var resources = new List<UnitRateResourceAmount>();
            decimal directMaterialBase = 0m;
            foreach (NormConsumption consumption in validation.NormResult.Consumptions)
            {
                if (consumption.Quantity == 0m)
                    continue;
                if (IsPercentage(consumption))
                    continue;
                UnitRateResourceBinding binding;
                bindings.TryGetValue(consumption.ResourceCode, out binding);
                string priceCode = binding?.PriceCode ?? consumption.ResourceCode;
                PriceProfilePrice price = request.PriceProfile.FindRequired(priceCode);
                decimal amount = consumption.Quantity * price.AppliedUnitPriceVnd;
                if (consumption.Kind == NormResourceKind.Material)
                    directMaterialBase += amount;
                resources.Add(CreateAmount(consumption, price, amount, binding));
            }

            foreach (NormConsumption consumption in validation.NormResult.Consumptions)
            {
                if (!IsPercentage(consumption))
                    continue;
                decimal amount = directMaterialBase * consumption.Quantity / 100m;
                resources.Add(new UnitRateResourceAmount(
                    consumption.Kind,
                    consumption.ResourceCode,
                    string.Empty,
                    "Vat lieu khac",
                    consumption.Unit,
                    consumption.Quantity,
                    0m,
                    amount,
                    true,
                    false,
                    FormatSource(request.Definition.Source),
                    string.Empty));
            }

            UnitRateResourceAmount[] ordered = resources
                .OrderBy(item => item.Kind)
                .ThenBy(item => item.ResourceCode, StringComparer.Ordinal)
                .ToArray();
            return new UnitRateCalculationResult(
                request,
                ordered,
                ordered.Where(item => item.Kind == NormResourceKind.Material).Sum(item => item.AmountVnd),
                ordered.Where(item => item.Kind == NormResourceKind.Labor).Sum(item => item.AmountVnd),
                ordered.Where(item => item.Kind == NormResourceKind.Machine).Sum(item => item.AmountVnd));
        }

        public static bool IsLogicalResource(string resourceCode)
        {
            string code = (resourceCode ?? string.Empty).Trim();
            return code.IndexOf("-OR-", StringComparison.OrdinalIgnoreCase) >= 0 ||
                code.EndsWith(".DIVING", StringComparison.OrdinalIgnoreCase);
        }

        private static Dictionary<string, UnitRateResourceBinding> ValidateBindings(
            UnitRateCalculationRequest request,
            NormCalculationResult normResult,
            ICollection<UnitRateValidationIssue> issues)
        {
            var result = new Dictionary<string, UnitRateResourceBinding>(StringComparer.OrdinalIgnoreCase);
            var resources = new HashSet<string>(
                normResult.Consumptions.Select(item => item.ResourceCode),
                StringComparer.OrdinalIgnoreCase);
            foreach (UnitRateResourceBinding binding in request.Bindings)
            {
                if (binding.ResourceCode.Length == 0 || binding.PriceCode.Length == 0 ||
                    binding.Reason.Length == 0 || binding.Reason.Length > 1024)
                {
                    issues.Add(new UnitRateValidationIssue(
                        UnitRateIssueCode.InvalidBinding,
                        binding.ResourceCode,
                        "Binding phai co ma nguon luc, ma gia va ly do 1-1024 ky tu."));
                    continue;
                }
                if (!resources.Contains(binding.ResourceCode))
                {
                    issues.Add(new UnitRateValidationIssue(
                        UnitRateIssueCode.InvalidBinding,
                        binding.ResourceCode,
                        "Binding khong thuoc dinh muc dang tinh: " + binding.ResourceCode + "."));
                    continue;
                }
                if (result.ContainsKey(binding.ResourceCode))
                {
                    issues.Add(new UnitRateValidationIssue(
                        UnitRateIssueCode.InvalidBinding,
                        binding.ResourceCode,
                        "Binding bi trung cho " + binding.ResourceCode + "."));
                    continue;
                }
                result.Add(binding.ResourceCode, binding);
            }
            return result;
        }

        private static UnitRateResourceAmount CreateAmount(
            NormConsumption consumption,
            PriceProfilePrice price,
            decimal amount,
            UnitRateResourceBinding binding)
        {
            return new UnitRateResourceAmount(
                consumption.Kind,
                consumption.ResourceCode,
                price.Entry.Code,
                price.Entry.DisplayName,
                consumption.Unit,
                consumption.Quantity,
                price.AppliedUnitPriceVnd,
                amount,
                false,
                price.IsOverridden,
                price.Override?.SourceReference ?? price.Entry.SourceReference,
                binding?.Reason ?? string.Empty);
        }

        private static bool IsPercentage(NormConsumption consumption)
        {
            return consumption.Kind == NormResourceKind.Material &&
                string.Equals(consumption.Unit, "percent", StringComparison.OrdinalIgnoreCase);
        }

        private static string FormatSource(RegulationSourceLocator source)
        {
            if (source == null)
                return string.Empty;
            string pages = source.PageFrom == source.PageTo
                ? source.PageFrom.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : source.PageFrom.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-" +
                    source.PageTo.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return source.DocumentId + ", page " + pages + "; " + source.Section;
        }

        private static PriceResourceKind ToPriceKind(NormResourceKind kind)
        {
            switch (kind)
            {
                case NormResourceKind.Material: return PriceResourceKind.Material;
                case NormResourceKind.Labor: return PriceResourceKind.Labor;
                case NormResourceKind.Machine: return PriceResourceKind.MachineShift;
                default: throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }
    }
}
