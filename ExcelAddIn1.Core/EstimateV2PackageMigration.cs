using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelAddIn1.Core
{
    public sealed class EstimateV2PackageMigrationPlan
    {
        private EstimateV2PackageMigrationPlan(EstimateV2State source, EstimateV2State target,
            RegulationPackage package, IEnumerable<string> errors, int affected)
        {
            SourcePayload = EstimateV2StateSerializer.Serialize(source);
            TargetState = target;
            Package = package;
            Errors = errors.ToList().AsReadOnly();
            AffectedCount = affected;
        }

        public string SourcePayload { get; }
        public EstimateV2State TargetState { get; }
        public RegulationPackage Package { get; }
        public IReadOnlyList<string> Errors { get; }
        public int AffectedCount { get; }
        public bool CanApply => Errors.Count == 0;

        public static EstimateV2PackageMigrationPlan Create(EstimateV2State source,
            RegulationPackageBundle bundle)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (bundle == null) throw new ArgumentNullException(nameof(bundle));
            RegulationDataModule module;
            if (!bundle.Modules.TryGetValue(RegulationModuleKind.Norm, out module))
                throw new InvalidOperationException("Gói dữ liệu không có module định mức.");
            NormCatalog catalog = NormCatalog.Load(module);
            var errors = new List<string>();
            var items = new List<EstimateV2WorkItemState>();
            int affected = 0;
            foreach (EstimateV2WorkItemState item in source.WorkItems)
            {
                if (!item.HasNormBinding || item.IsOrphaned)
                {
                    items.Add(item);
                    continue;
                }
                try
                {
                    NormDefinition norm = catalog.FindRequired(item.NormCode);
                    if (!norm.Variants.Contains(item.VariantCode, StringComparer.Ordinal))
                        throw new InvalidOperationException("Thiếu variant " + item.VariantCode);
                    if (!string.Equals(item.PackageId, bundle.Package.PackageId, StringComparison.Ordinal) ||
                        !string.Equals(item.DataVersion, bundle.Package.DataVersion, StringComparison.Ordinal) ||
                        !string.Equals(item.PackageChecksum, bundle.Package.PackageChecksum, StringComparison.OrdinalIgnoreCase))
                        affected++;
                    items.Add(item.WithBinding(item.NormCode, item.VariantCode, bundle.Package.PackageId,
                        bundle.Package.DataVersion, bundle.Package.PackageChecksum));
                }
                catch (Exception ex) when (ex is KeyNotFoundException || ex is InvalidOperationException)
                {
                    errors.Add(item.WorkItemId + ": " + item.NormCode + "/" + item.VariantCode + ": " + ex.Message);
                    items.Add(item);
                }
            }
            return new EstimateV2PackageMigrationPlan(source,
                new EstimateV2State(items, source.UpdatedUtc), bundle.Package, errors, affected);
        }

        public bool Matches(EstimateV2State current)
        {
            return current != null && string.Equals(SourcePayload,
                EstimateV2StateSerializer.Serialize(current), StringComparison.Ordinal);
        }
    }
}
