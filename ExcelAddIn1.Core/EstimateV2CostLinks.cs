using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public enum EstimateV2CostLinkStatus
    {
        Unbound = 0,
        RateNotResolved = 1,
        Ready = 2
    }

    public sealed class EstimateV2WorkItemCostLink
    {
        internal EstimateV2WorkItemCostLink(
            string workItemId,
            string sourceKey,
            string normCode,
            string variantCode,
            string packageIdentity,
            string rateId,
            EstimateV2CostLinkStatus status,
            bool requiresConditionReview)
        {
            WorkItemId = (workItemId ?? string.Empty).Trim();
            SourceKey = (sourceKey ?? string.Empty).Trim();
            NormCode = (normCode ?? string.Empty).Trim();
            VariantCode = (variantCode ?? string.Empty).Trim();
            PackageIdentity = (packageIdentity ?? string.Empty).Trim();
            RateId = (rateId ?? string.Empty).Trim();
            Status = status;
            RequiresConditionReview = requiresConditionReview;
        }

        public string WorkItemId { get; }
        public string SourceKey { get; }
        public string NormCode { get; }
        public string VariantCode { get; }
        public string PackageIdentity { get; }
        public string RateId { get; }
        public EstimateV2CostLinkStatus Status { get; }
        public bool RequiresConditionReview { get; }

        public bool IsReady =>
            Status == EstimateV2CostLinkStatus.Ready &&
            RateId.Length > 0;
    }

    public sealed class EstimateV2CostLinkPlan
    {
        private readonly IReadOnlyDictionary<string, EstimateV2WorkItemCostLink> byWorkItemId;

        private EstimateV2CostLinkPlan(
            IEnumerable<EstimateV2WorkItemCostLink> links)
        {
            EstimateV2WorkItemCostLink[] ordered = (links ??
                Enumerable.Empty<EstimateV2WorkItemCostLink>())
                .Where(item => item != null)
                .OrderBy(item => item.SourceKey, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.WorkItemId, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            Links = new ReadOnlyCollection<EstimateV2WorkItemCostLink>(ordered);
            byWorkItemId = new ReadOnlyDictionary<string, EstimateV2WorkItemCostLink>(
                ordered.ToDictionary(
                    item => item.WorkItemId,
                    StringComparer.OrdinalIgnoreCase));
        }

        public IReadOnlyList<EstimateV2WorkItemCostLink> Links { get; }
        public int TotalCount => Links.Count;
        public int UnboundCount => Links.Count(item =>
            item.Status == EstimateV2CostLinkStatus.Unbound);
        public int MissingRateCount => Links.Count(item =>
            item.Status == EstimateV2CostLinkStatus.RateNotResolved);
        public int ReadyCount => Links.Count(item => item.IsReady);
        public int ConditionReviewCount => Links.Count(item =>
            item.IsReady && item.RequiresConditionReview);

        public EstimateV2WorkItemCostLink Find(string workItemId)
        {
            EstimateV2WorkItemCostLink value;
            return byWorkItemId.TryGetValue(
                (workItemId ?? string.Empty).Trim(),
                out value)
                ? value
                : null;
        }

        public static EstimateV2CostLinkPlan Build(
            EstimateV2State state,
            EstimateV2RatePlan ratePlan)
        {
            if (state == null)
                throw new ArgumentNullException(nameof(state));
            if (ratePlan == null)
                throw new ArgumentNullException(nameof(ratePlan));

            var links = new List<EstimateV2WorkItemCostLink>();
            foreach (EstimateV2WorkItemState item in state.WorkItems
                .Where(value => value != null && !value.IsOrphaned))
            {
                string packageIdentity =
                    EstimateV2ResourcePlanBuilder.PackageIdentity(item);

                if (!item.HasNormBinding)
                {
                    links.Add(new EstimateV2WorkItemCostLink(
                        item.WorkItemId,
                        item.SourceKey,
                        item.NormCode,
                        item.VariantCode,
                        packageIdentity,
                        string.Empty,
                        EstimateV2CostLinkStatus.Unbound,
                        false));
                    continue;
                }

                EstimateV2RateItem[] candidates = ratePlan.Items
                    .Where(rate =>
                        string.Equals(
                            rate.PackageIdentity,
                            packageIdentity,
                            StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(
                            rate.NormCode,
                            item.NormCode,
                            StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                EstimateV2RateItem resolved = null;
                string wantedVariant =
                    (item.VariantCode ?? string.Empty).Trim();
                if (wantedVariant.Length > 0)
                {
                    resolved = candidates.SingleOrDefault(rate =>
                        string.Equals(
                            rate.VariantCode,
                            wantedVariant,
                            StringComparison.OrdinalIgnoreCase));
                }
                else if (candidates.Length == 1)
                {
                    // Binding cu co the chua luu variant khi dinh muc chi co mot variant.
                    resolved = candidates[0];
                }

                if (resolved == null)
                {
                    links.Add(new EstimateV2WorkItemCostLink(
                        item.WorkItemId,
                        item.SourceKey,
                        item.NormCode,
                        item.VariantCode,
                        packageIdentity,
                        string.Empty,
                        EstimateV2CostLinkStatus.RateNotResolved,
                        false));
                    continue;
                }

                links.Add(new EstimateV2WorkItemCostLink(
                    item.WorkItemId,
                    item.SourceKey,
                    resolved.NormCode,
                    resolved.VariantCode,
                    packageIdentity,
                    resolved.RateId,
                    EstimateV2CostLinkStatus.Ready,
                    resolved.RequiresConditionReview));
            }

            return new EstimateV2CostLinkPlan(links);
        }
    }
}
