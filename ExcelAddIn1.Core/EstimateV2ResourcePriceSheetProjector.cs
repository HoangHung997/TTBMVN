using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    /// <summary>
    /// Projects logical norm resources into the concrete resources that must
    /// have a price available on the VL-NC-M sheet. This does not decide which
    /// alternative will later be used by a unit-rate block; it only makes sure
    /// all eligible concrete prices can be calculated/entered in Excel.
    /// </summary>
    public static class EstimateV2ResourcePriceSheetProjector
    {
        private static readonly string[] DivingMachineCandidates =
        {
            "M010.029",
            "M010.030",
            "M010.031",
            "M010.032",
            "M010.033"
        };

        public static IReadOnlyList<EstimateV2ResourceRequirement> Project(
            EstimateV2ResourcePlan plan)
        {
            if (plan == null)
                throw new ArgumentNullException(nameof(plan));

            var merged = new Dictionary<string, MutableRequirement>(
                StringComparer.OrdinalIgnoreCase);

            foreach (EstimateV2ResourceRequirement item in plan.Resources)
            {
                string[] candidates = ExpandCodes(item).ToArray();
                if (candidates.Length == 0)
                    candidates = new[] { item.Code };

                foreach (string code in candidates)
                    Add(merged, item, code);
            }

            return new ReadOnlyCollection<EstimateV2ResourceRequirement>(
                merged.Values
                    .Select(item => item.Freeze())
                    .OrderBy(item => item.Kind)
                    .ThenBy(item => item.Code, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(item => item.Unit, StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        private static IEnumerable<string> ExpandCodes(
            EstimateV2ResourceRequirement item)
        {
            if (item == null)
                yield break;

            string code = (item.Code ?? string.Empty).Trim();
            if (!item.IsLogical || code.Length == 0)
            {
                yield return code;
                yield break;
            }

            if (item.Kind == NormResourceKind.Machine &&
                code.EndsWith(".DIVING", StringComparison.OrdinalIgnoreCase))
            {
                foreach (string candidate in DivingMachineCandidates)
                    yield return candidate;
                yield break;
            }

            const string separator = "-OR-";
            int split = code.IndexOf(separator, StringComparison.OrdinalIgnoreCase);
            if (split > 0)
            {
                string first = code.Substring(0, split).Trim();
                string second = code.Substring(split + separator.Length).Trim();
                if (first.Length > 0)
                    yield return first;
                if (second.Length > 0)
                    yield return CompleteAlternative(first, second);
                yield break;
            }

            // Unknown logical convention: retain the code so the UI can still
            // expose an input instead of making the whole VL-NC-M step unusable.
            yield return code;
        }

        private static string CompleteAlternative(
            string first,
            string second)
        {
            if (string.IsNullOrWhiteSpace(second))
                return string.Empty;

            if (second.IndexOf('-', StringComparison.Ordinal) >= 0 ||
                second.IndexOf('.', StringComparison.Ordinal) >= 0)
            {
                return second;
            }

            int prefixEnd = (first ?? string.Empty).IndexOf(
                '-',
                StringComparison.Ordinal);
            if (prefixEnd >= 0)
                return first.Substring(0, prefixEnd + 1) + second;
            return second;
        }

        private static void Add(
            IDictionary<string, MutableRequirement> merged,
            EstimateV2ResourceRequirement source,
            string code)
        {
            string normalized = (code ?? string.Empty).Trim();
            if (normalized.Length == 0)
                return;

            string key = ((int)source.Kind) + "|" +
                normalized.ToUpperInvariant() + "|" +
                (source.Unit ?? string.Empty).Trim().ToUpperInvariant();

            MutableRequirement current;
            if (!merged.TryGetValue(key, out current))
            {
                current = new MutableRequirement(
                    source.Kind,
                    normalized,
                    source.Unit,
                    source.RequiresUnitPrice);
                merged.Add(key, current);
            }

            current.UsageCount += source.UsageCount;
            foreach (string norm in source.NormKeys)
                current.NormKeys.Add(norm);
            foreach (string package in source.PackageIdentities)
                current.PackageIdentities.Add(package);
        }

        private sealed class MutableRequirement
        {
            internal MutableRequirement(
                NormResourceKind kind,
                string code,
                string unit,
                bool requiresUnitPrice)
            {
                Kind = kind;
                Code = code;
                Unit = unit ?? string.Empty;
                RequiresUnitPrice = requiresUnitPrice;
                NormKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                PackageIdentities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            internal NormResourceKind Kind;
            internal string Code;
            internal string Unit;
            internal bool RequiresUnitPrice;
            internal int UsageCount;
            internal HashSet<string> NormKeys;
            internal HashSet<string> PackageIdentities;

            internal EstimateV2ResourceRequirement Freeze()
            {
                return new EstimateV2ResourceRequirement(
                    Kind,
                    Code,
                    Unit,
                    RequiresUnitPrice,
                    false,
                    UsageCount,
                    NormKeys,
                    PackageIdentities);
            }
        }
    }
}
