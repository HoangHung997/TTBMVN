using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public sealed class EstimateV2ResourceRequirement
    {
        internal EstimateV2ResourceRequirement(
            NormResourceKind kind,
            string code,
            string unit,
            bool requiresUnitPrice,
            bool isLogical,
            int usageCount,
            IEnumerable<string> normKeys,
            IEnumerable<string> packageIdentities)
        {
            Kind = kind;
            Code = (code ?? string.Empty).Trim();
            Unit = (unit ?? string.Empty).Trim();
            RequiresUnitPrice = requiresUnitPrice;
            IsLogical = isLogical;
            UsageCount = usageCount;
            NormKeys = new ReadOnlyCollection<string>(
                (normKeys ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
            PackageIdentities = new ReadOnlyCollection<string>(
                (packageIdentities ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public NormResourceKind Kind { get; }
        public string Code { get; }
        public string Unit { get; }
        public bool RequiresUnitPrice { get; }
        public bool IsLogical { get; }
        public int UsageCount { get; }
        public IReadOnlyList<string> NormKeys { get; }
        public IReadOnlyList<string> PackageIdentities { get; }
    }

    public sealed class EstimateV2ResourcePlan
    {
        public EstimateV2ResourcePlan(
            IEnumerable<EstimateV2ResourceRequirement> resources,
            int boundWorkItemCount,
            int uniqueNormBindingCount)
        {
            Resources = new ReadOnlyCollection<EstimateV2ResourceRequirement>(
                (resources ?? Enumerable.Empty<EstimateV2ResourceRequirement>())
                    .OrderBy(item => item.Kind)
                    .ThenBy(item => item.Code, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(item => item.Unit, StringComparer.OrdinalIgnoreCase)
                    .ToList());
            BoundWorkItemCount = boundWorkItemCount;
            UniqueNormBindingCount = uniqueNormBindingCount;
        }

        public IReadOnlyList<EstimateV2ResourceRequirement> Resources { get; }
        public int BoundWorkItemCount { get; }
        public int UniqueNormBindingCount { get; }

        public IReadOnlyList<EstimateV2ResourceRequirement> Materials =>
            Filter(NormResourceKind.Material);

        public IReadOnlyList<EstimateV2ResourceRequirement> Labor =>
            Filter(NormResourceKind.Labor);

        public IReadOnlyList<EstimateV2ResourceRequirement> Machines =>
            Filter(NormResourceKind.Machine);

        public IReadOnlyList<EstimateV2ResourceRequirement> LogicalResources =>
            new ReadOnlyCollection<EstimateV2ResourceRequirement>(
                Resources.Where(item => item.IsLogical).ToList());

        private IReadOnlyList<EstimateV2ResourceRequirement> Filter(
            NormResourceKind kind)
        {
            return new ReadOnlyCollection<EstimateV2ResourceRequirement>(
                Resources.Where(item => item.Kind == kind).ToList());
        }
    }

    public static class EstimateV2ResourcePlanBuilder
    {
        public static EstimateV2ResourcePlan Build(
            IEnumerable<EstimateV2WorkItemState> workItems,
            Func<EstimateV2WorkItemState, NormDefinition> definitionResolver)
        {
            if (definitionResolver == null)
                throw new ArgumentNullException(nameof(definitionResolver));

            EstimateV2WorkItemState[] active = (workItems ??
                Enumerable.Empty<EstimateV2WorkItemState>())
                .Where(item => item != null && !item.IsOrphaned && item.HasNormBinding)
                .ToArray();

            var resources = new Dictionary<string, MutableRequirement>(
                StringComparer.OrdinalIgnoreCase);
            var uniqueNorms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (EstimateV2WorkItemState item in active)
            {
                NormDefinition definition = definitionResolver(item);
                if (definition == null)
                {
                    throw new InvalidOperationException(
                        "Khong resolve duoc dinh muc " + item.NormCode + ".");
                }
                if (!string.Equals(
                    definition.Key,
                    item.NormCode,
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Norm resolver tra sai key: can " + item.NormCode +
                        ", co " + definition.Key + ".");
                }

                int variantIndex = ResolveVariantIndex(definition, item.VariantCode);
                string normBinding = definition.Key + "|" +
                    (item.VariantCode ?? string.Empty).Trim();
                uniqueNorms.Add(PackageIdentity(item) + "|" + normBinding);

                foreach (NormResourceRate rate in definition.Rates)
                {
                    if (rate == null || rate.Quantities.Count <= variantIndex)
                        continue;
                    decimal quantity = rate.Quantities[variantIndex];
                    if (quantity == 0m)
                        continue;

                    bool percentage = rate.Kind == NormResourceKind.Material &&
                        string.Equals(
                            rate.Unit,
                            "percent",
                            StringComparison.OrdinalIgnoreCase);
                    string key = ((int)rate.Kind) + "|" +
                        (rate.ResourceCode ?? string.Empty).Trim() + "|" +
                        (rate.Unit ?? string.Empty).Trim();

                    MutableRequirement current;
                    if (!resources.TryGetValue(key, out current))
                    {
                        current = new MutableRequirement(
                            rate.Kind,
                            rate.ResourceCode,
                            rate.Unit,
                            !percentage,
                            UnitRateCalculator.IsLogicalResource(rate.ResourceCode));
                        resources.Add(key, current);
                    }
                    current.UsageCount++;
                    current.NormKeys.Add(normBinding);
                    current.PackageIdentities.Add(PackageIdentity(item));
                }
            }

            return new EstimateV2ResourcePlan(
                resources.Values.Select(item => item.Freeze()),
                active.Length,
                uniqueNorms.Count);
        }

        private static int ResolveVariantIndex(
            NormDefinition definition,
            string variantCode)
        {
            if (definition.Variants == null || definition.Variants.Count == 0)
                return 0;

            string wanted = (variantCode ?? string.Empty).Trim();
            if (wanted.Length == 0)
            {
                if (definition.Variants.Count == 1)
                    return 0;
                throw new InvalidOperationException(
                    definition.Key + ": chua chon variant dinh muc.");
            }

            for (int index = 0; index < definition.Variants.Count; index++)
            {
                if (string.Equals(
                    definition.Variants[index],
                    wanted,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return index;
                }
            }
            throw new KeyNotFoundException(
                definition.Key + ": khong co variant " + wanted + ".");
        }

        public static string PackageIdentity(EstimateV2WorkItemState item)
        {
            if (item == null)
                return string.Empty;
            return (item.PackageId ?? string.Empty).Trim() + "@" +
                (item.DataVersion ?? string.Empty).Trim() + "#" +
                (item.PackageChecksum ?? string.Empty).Trim().ToUpperInvariant();
        }

        private sealed class MutableRequirement
        {
            internal MutableRequirement(
                NormResourceKind kind,
                string code,
                string unit,
                bool requiresUnitPrice,
                bool isLogical)
            {
                Kind = kind;
                Code = (code ?? string.Empty).Trim();
                Unit = (unit ?? string.Empty).Trim();
                RequiresUnitPrice = requiresUnitPrice;
                IsLogical = isLogical;
                NormKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                PackageIdentities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            internal NormResourceKind Kind { get; }
            internal string Code { get; }
            internal string Unit { get; }
            internal bool RequiresUnitPrice { get; }
            internal bool IsLogical { get; }
            internal int UsageCount { get; set; }
            internal ISet<string> NormKeys { get; }
            internal ISet<string> PackageIdentities { get; }

            internal EstimateV2ResourceRequirement Freeze()
            {
                return new EstimateV2ResourceRequirement(
                    Kind,
                    Code,
                    Unit,
                    RequiresUnitPrice,
                    IsLogical,
                    UsageCount,
                    NormKeys,
                    PackageIdentities);
            }
        }
    }
}
