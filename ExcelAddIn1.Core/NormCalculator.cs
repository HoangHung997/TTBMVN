using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public sealed class NormConsumption
    {
        internal NormConsumption(
            NormResourceKind kind,
            string resourceCode,
            string unit,
            decimal quantity)
        {
            Kind = kind;
            ResourceCode = resourceCode;
            Unit = unit;
            Quantity = quantity;
        }

        public NormResourceKind Kind { get; }
        public string ResourceCode { get; }
        public string Unit { get; }
        public decimal Quantity { get; }
    }

    public sealed class NormCalculationResult
    {
        internal NormCalculationResult(
            string normKey,
            string variantCode,
            decimal workQuantity,
            IEnumerable<NormConsumption> consumptions,
            RegulationSourceLocator source)
        {
            NormKey = normKey;
            VariantCode = variantCode;
            WorkQuantity = workQuantity;
            Consumptions = new ReadOnlyCollection<NormConsumption>(consumptions.ToList());
            Source = source;
        }

        public string NormKey { get; }
        public string VariantCode { get; }
        public decimal WorkQuantity { get; }
        public IReadOnlyList<NormConsumption> Consumptions { get; }
        public RegulationSourceLocator Source { get; }
        public string RoundingRule => "none-invariant-decimal";

        public NormConsumption FindRequired(NormResourceKind kind, string resourceCode)
        {
            NormConsumption value = Consumptions.SingleOrDefault(
                item => item.Kind == kind &&
                    string.Equals(item.ResourceCode, resourceCode, StringComparison.Ordinal));
            if (value == null)
                throw new KeyNotFoundException("Khong co hao phi: " + kind + ":" + resourceCode + ".");
            return value;
        }
    }

    public static class NormCalculator
    {
        public static NormCalculationResult Calculate(
            NormDefinition definition,
            string variantCode,
            decimal workQuantity,
            IEnumerable<string> conditions = null)
        {
            if (definition == null)
                throw new ArgumentNullException(nameof(definition));
            if (workQuantity < 0m)
                throw new ArgumentOutOfRangeException(nameof(workQuantity), "Khoi luong phai khong am.");

            string variant = (variantCode ?? string.Empty).Trim();
            int variantIndex = definition.Variants.IndexOf(variant);
            if (variantIndex < 0)
                throw new ArgumentException("Bien the dinh muc khong ton tai: " + variantCode + ".", nameof(variantCode));

            var selected = new HashSet<string>(
                (conditions ?? Enumerable.Empty<string>())
                    .Select(value => (value ?? string.Empty).Trim())
                    .Where(value => value.Length > 0),
                StringComparer.Ordinal);
            ValidateConditions(definition, variant, selected);

            var values = definition.Rates.ToDictionary(
                rate => Identity(rate.Kind, rate.ResourceCode),
                rate => new MutableConsumption(
                    rate.Kind,
                    rate.ResourceCode,
                    rate.Unit,
                    rate.Quantities[variantIndex]),
                StringComparer.Ordinal);

            foreach (NormAdjustment adjustment in definition.Adjustments)
            {
                if (!selected.Contains(adjustment.Condition))
                    continue;
                if (adjustment.Operation == NormAdjustmentOperation.Factor)
                {
                    foreach (MutableConsumption value in values.Values)
                    {
                        if (adjustment.TargetKinds.Contains(value.Kind))
                            value.Quantity *= adjustment.Value;
                    }
                }
                else
                {
                    NormResourceKind kind = adjustment.TargetKinds.Single();
                    string identity = Identity(kind, adjustment.ResourceCode);
                    if (values.TryGetValue(identity, out MutableConsumption existing))
                    {
                        if (!string.Equals(existing.Unit, adjustment.Unit, StringComparison.Ordinal))
                            throw new InvalidOperationException("Don vi adjustment khong khop resource " + identity + ".");
                        existing.Quantity += adjustment.Value;
                    }
                    else
                    {
                        values.Add(identity, new MutableConsumption(
                            kind,
                            adjustment.ResourceCode,
                            adjustment.Unit,
                            adjustment.Value));
                    }
                }
            }

            NormConsumption[] consumptions = values.Values
                .OrderBy(value => value.Kind)
                .ThenBy(value => value.ResourceCode, StringComparer.Ordinal)
                .Select(value => new NormConsumption(
                    value.Kind,
                    value.ResourceCode,
                    value.Unit,
                    value.Quantity * workQuantity))
                .ToArray();
            return new NormCalculationResult(
                definition.Key,
                variant,
                workQuantity,
                consumptions,
                definition.Source);
        }

        private static void ValidateConditions(
            NormDefinition definition,
            string variant,
            ISet<string> selected)
        {
            var known = new HashSet<string>(
                definition.Adjustments.Select(item => item.Condition),
                StringComparer.Ordinal);
            foreach (NormConstraint constraint in definition.Constraints)
            {
                if (constraint.Operation == "prohibit-condition")
                    known.Add(constraint.Argument1);
                else if (constraint.Operation == "requires-condition")
                    known.Add(constraint.Argument2);
            }
            foreach (string condition in selected)
            {
                if (!known.Contains(condition))
                    throw new ArgumentException("Dieu kien khong duoc ho tro: " + condition + ".", nameof(selected));
            }

            if (selected.Count(value => value.StartsWith("current-", StringComparison.Ordinal)) > 1)
                throw new ArgumentException("Chi duoc chon mot khoang van toc dong chay.", nameof(selected));

            foreach (NormConstraint constraint in definition.Constraints)
            {
                if (constraint.Operation == "prohibit-condition" && selected.Contains(constraint.Argument1))
                {
                    throw new InvalidOperationException(
                        definition.Key + " khong cho phep dieu kien " + constraint.Argument1 + ".");
                }
                if (constraint.Operation == "requires-condition" &&
                    string.Equals(variant, constraint.Argument1, StringComparison.Ordinal) &&
                    !selected.Contains(constraint.Argument2))
                {
                    throw new InvalidOperationException(
                        variant + " yeu cau dieu kien " + constraint.Argument2 + ".");
                }
            }
        }

        private static string Identity(NormResourceKind kind, string resourceCode)
        {
            return kind + ":" + resourceCode;
        }

        private sealed class MutableConsumption
        {
            internal MutableConsumption(
                NormResourceKind kind,
                string resourceCode,
                string unit,
                decimal quantity)
            {
                Kind = kind;
                ResourceCode = resourceCode;
                Unit = unit;
                Quantity = quantity;
            }

            internal NormResourceKind Kind { get; }
            internal string ResourceCode { get; }
            internal string Unit { get; }
            internal decimal Quantity { get; set; }
        }
    }

    internal static class ReadOnlyListExtensions
    {
        internal static int IndexOf<T>(this IReadOnlyList<T> values, T value)
        {
            for (int index = 0; index < values.Count; index++)
            {
                if (EqualityComparer<T>.Default.Equals(values[index], value))
                    return index;
            }
            return -1;
        }
    }
}
