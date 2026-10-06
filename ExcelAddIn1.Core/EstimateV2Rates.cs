using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ExcelAddIn1.Core
{
    public enum EstimateV2RateEnvironment
    {
        Land = 1,
        InlandWater = 2,
        Sea = 3
    }

    public sealed class EstimateV2RateResource
    {
        public EstimateV2RateResource(
            NormResourceKind kind,
            string resourceCode,
            string unit,
            decimal quantity,
            bool isPercentage,
            IEnumerable<string> priceCandidates)
        {
            Kind = kind;
            ResourceCode = Clean(resourceCode);
            Unit = Clean(unit);
            Quantity = quantity;
            IsPercentage = isPercentage;
            PriceCandidates = new ReadOnlyCollection<string>(
                (priceCandidates ?? Enumerable.Empty<string>())
                    .Select(Clean)
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public NormResourceKind Kind { get; }
        public string ResourceCode { get; }
        public string Unit { get; }
        public decimal Quantity { get; }
        public bool IsPercentage { get; }
        public IReadOnlyList<string> PriceCandidates { get; }

        private static string Clean(string value)
        {
            return (value ?? string.Empty).Trim();
        }
    }

    public sealed class EstimateV2RateItem
    {
        public EstimateV2RateItem(
            string rateId,
            string packageIdentity,
            string normCode,
            string variantCode,
            string title,
            string workUnit,
            EstimateV2RateEnvironment environment,
            int usageCount,
            bool requiresConditionReview,
            IEnumerable<EstimateV2RateResource> resources)
        {
            RateId = Required(rateId, nameof(rateId));
            PackageIdentity = Required(packageIdentity, nameof(packageIdentity));
            NormCode = Required(normCode, nameof(normCode));
            VariantCode = Required(variantCode, nameof(variantCode));
            Title = (title ?? string.Empty).Trim();
            WorkUnit = (workUnit ?? string.Empty).Trim();
            Environment = environment;
            UsageCount = usageCount;
            RequiresConditionReview = requiresConditionReview;
            Resources = new ReadOnlyCollection<EstimateV2RateResource>(
                (resources ?? Enumerable.Empty<EstimateV2RateResource>())
                    .Where(item => item != null && item.Quantity != 0m)
                    .OrderBy(item => item.Kind)
                    .ThenBy(item => item.IsPercentage)
                    .ThenBy(item => item.ResourceCode, StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public string RateId { get; }
        public string PackageIdentity { get; }
        public string NormCode { get; }
        public string VariantCode { get; }
        public string Title { get; }
        public string WorkUnit { get; }
        public EstimateV2RateEnvironment Environment { get; }
        public int UsageCount { get; }
        public bool RequiresConditionReview { get; }
        public IReadOnlyList<EstimateV2RateResource> Resources { get; }

        public IReadOnlyList<EstimateV2RateResource> Materials =>
            Filter(NormResourceKind.Material);

        public IReadOnlyList<EstimateV2RateResource> Labor =>
            Filter(NormResourceKind.Labor);

        public IReadOnlyList<EstimateV2RateResource> Machines =>
            Filter(NormResourceKind.Machine);

        private IReadOnlyList<EstimateV2RateResource> Filter(NormResourceKind kind)
        {
            return new ReadOnlyCollection<EstimateV2RateResource>(
                Resources.Where(item => item.Kind == kind).ToList());
        }

        private static string Required(string value, string name)
        {
            string text = (value ?? string.Empty).Trim();
            if (text.Length == 0)
                throw new ArgumentException(name + " khong duoc trong.", name);
            return text;
        }
    }

    public sealed class EstimateV2RatePlan
    {
        public EstimateV2RatePlan(IEnumerable<EstimateV2RateItem> items)
        {
            Items = new ReadOnlyCollection<EstimateV2RateItem>(
                (items ?? Enumerable.Empty<EstimateV2RateItem>())
                    .Where(item => item != null)
                    .OrderBy(item => item.Environment)
                    .ThenBy(item => item.NormCode, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(item => item.VariantCode, StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public IReadOnlyList<EstimateV2RateItem> Items { get; }

        public IReadOnlyList<EstimateV2RateItem> ForEnvironment(
            EstimateV2RateEnvironment environment)
        {
            return new ReadOnlyCollection<EstimateV2RateItem>(
                Items.Where(item => item.Environment == environment).ToList());
        }
    }

    public static class EstimateV2RatePlanBuilder
    {
        public static EstimateV2RatePlan Build(
            IEnumerable<EstimateV2WorkItemState> workItems,
            Func<EstimateV2WorkItemState, NormDefinition> definitionResolver)
        {
            if (definitionResolver == null)
                throw new ArgumentNullException(nameof(definitionResolver));

            EstimateV2WorkItemState[] active = (workItems ??
                Enumerable.Empty<EstimateV2WorkItemState>())
                .Where(item => item != null && !item.IsOrphaned && item.HasNormBinding)
                .ToArray();

            var items = new List<EstimateV2RateItem>();
            foreach (IGrouping<string, EstimateV2WorkItemState> group in active.GroupBy(
                item => CanonicalRateKey(item),
                StringComparer.OrdinalIgnoreCase))
            {
                EstimateV2WorkItemState sample = group.First();
                NormDefinition definition = definitionResolver(sample);
                if (definition == null)
                {
                    throw new InvalidOperationException(
                        "Khong resolve duoc dinh muc " + sample.NormCode + ".");
                }
                if (!string.Equals(
                    definition.Key,
                    sample.NormCode,
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Norm resolver tra sai key: can " + sample.NormCode +
                        ", co " + definition.Key + ".");
                }

                int variantIndex = ResolveVariantIndex(
                    definition,
                    sample.VariantCode);
                var resources = new List<EstimateV2RateResource>();
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
                    bool logical = UnitRateCalculator.IsLogicalResource(
                        rate.ResourceCode);
                    IReadOnlyList<string> candidates =
                        EstimateV2ResourcePriceSheetProjector.ExpandResourceCodes(
                            rate.Kind,
                            rate.ResourceCode,
                            logical);

                    resources.Add(new EstimateV2RateResource(
                        rate.Kind,
                        rate.ResourceCode,
                        rate.Unit,
                        quantity,
                        percentage,
                        percentage
                            ? Enumerable.Empty<string>()
                            : candidates));
                }

                string packageIdentity =
                    EstimateV2ResourcePlanBuilder.PackageIdentity(sample);
                items.Add(new EstimateV2RateItem(
                    CreateRateId(
                        packageIdentity,
                        definition.Key,
                        sample.VariantCode),
                    packageIdentity,
                    definition.Key,
                    sample.VariantCode,
                    definition.Title,
                    definition.WorkUnit,
                    ClassifyEnvironment(definition.Key),
                    group.Count(),
                    definition.Adjustments.Count > 0 ||
                        definition.Constraints.Count > 0,
                    resources));
            }

            return new EstimateV2RatePlan(items);
        }

        public static EstimateV2RateEnvironment ClassifyEnvironment(
            string normCode)
        {
            string code = (normCode ?? string.Empty).Trim().ToUpperInvariant();
            if (code.StartsWith("NORM-000.", StringComparison.Ordinal) ||
                code.StartsWith("NORM-010.", StringComparison.Ordinal) ||
                code.StartsWith("NORM-020.", StringComparison.Ordinal))
            {
                return EstimateV2RateEnvironment.Land;
            }
            if (code.StartsWith("NORM-030.", StringComparison.Ordinal))
                return EstimateV2RateEnvironment.InlandWater;
            if (code.StartsWith("NORM-040.", StringComparison.Ordinal))
                return EstimateV2RateEnvironment.Sea;

            throw new FormatException(
                "Khong phan loai duoc moi truong don gia cho " + normCode + ".");
        }

        public static string CreateRateId(
            string packageIdentity,
            string normCode,
            string variantCode)
        {
            string canonical = string.Join("|", new[]
            {
                (packageIdentity ?? string.Empty).Trim().ToUpperInvariant(),
                (normCode ?? string.Empty).Trim().ToUpperInvariant(),
                (variantCode ?? string.Empty).Trim().ToUpperInvariant()
            });
            using (SHA256 algorithm = SHA256.Create())
            {
                byte[] hash = algorithm.ComputeHash(
                    Encoding.UTF8.GetBytes(canonical));
                string hex = string.Concat(hash.Select(item =>
                    item.ToString("X2", CultureInfo.InvariantCulture)));
                return "DG-" + hex.Substring(0, 18);
            }
        }

        private static string CanonicalRateKey(EstimateV2WorkItemState item)
        {
            return string.Join("|", new[]
            {
                EstimateV2ResourcePlanBuilder.PackageIdentity(item),
                (item.NormCode ?? string.Empty).Trim(),
                (item.VariantCode ?? string.Empty).Trim()
            });
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
    }
}
