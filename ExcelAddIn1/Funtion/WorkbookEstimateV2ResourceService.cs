using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookEstimateV2ResourcePreview
    {
        internal WorkbookEstimateV2ResourcePreview(
            EstimateV2ResourcePlan plan,
            IEnumerable<string> missingPackageBindings,
            IEnumerable<string> unresolvedLogicalResources)
        {
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            PriceSheetResources = EstimateV2ResourcePriceSheetProjector.Project(Plan);
            MissingPackageBindings = new ReadOnlyCollection<string>(
                (missingPackageBindings ?? Enumerable.Empty<string>())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
            UnresolvedLogicalResources = new ReadOnlyCollection<string>(
                (unresolvedLogicalResources ?? Enumerable.Empty<string>())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public EstimateV2ResourcePlan Plan { get; }
        public IReadOnlyList<EstimateV2ResourceRequirement> PriceSheetResources { get; }
        public IReadOnlyList<string> MissingPackageBindings { get; }
        public IReadOnlyList<string> UnresolvedLogicalResources { get; }

        // Tai nguyen logic (OR/DIVING) khong chan VL-NC-M: sheet gia se sinh
        // tat ca ung vien vat ly. Viec chon ung vien cu the thuoc buoc don gia.
        public bool CanGenerate =>
            MissingPackageBindings.Count == 0 &&
            Plan.BoundWorkItemCount > 0;
    }

    public static class WorkbookEstimateV2ResourceService
    {
        public static WorkbookEstimateV2ResourcePreview BuildPreview(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            EstimateV2State state;
            if (!WorkbookEstimateV2StateService.TryLoad(workbook, out state))
            {
                return new WorkbookEstimateV2ResourcePreview(
                    new EstimateV2ResourcePlan(
                        Enumerable.Empty<EstimateV2ResourceRequirement>(),
                        0,
                        0),
                    Enumerable.Empty<string>(),
                    Enumerable.Empty<string>());
            }

            EstimateV2WorkItemState[] bound = state.WorkItems
                .Where(item => item != null && !item.IsOrphaned && item.HasNormBinding)
                .ToArray();
            if (bound.Length == 0)
            {
                return new WorkbookEstimateV2ResourcePreview(
                    new EstimateV2ResourcePlan(
                        Enumerable.Empty<EstimateV2ResourceRequirement>(),
                        0,
                        0),
                    Enumerable.Empty<string>(),
                    Enumerable.Empty<string>());
            }

            RegulationPackageBootstrapService.LoadAvailablePackages();
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            var catalogs = new Dictionary<string, NormCatalog>(
                StringComparer.OrdinalIgnoreCase);
            var missing = new List<string>();

            foreach (IGrouping<string, EstimateV2WorkItemState> packageGroup in bound
                .GroupBy(
                    EstimateV2ResourcePlanBuilder.PackageIdentity,
                    StringComparer.OrdinalIgnoreCase))
            {
                EstimateV2WorkItemState sample = packageGroup.First();
                if (sample.PackageId.Length == 0 ||
                    sample.DataVersion.Length == 0 ||
                    sample.PackageChecksum.Length == 0)
                {
                    missing.Add(
                        sample.NormCode +
                        ": binding cu/chua pin day du package. Hay gan lai dinh muc.");
                    continue;
                }

                try
                {
                    RegulationPackageBundle bundle = store.LoadBundleRequired(
                        sample.PackageId,
                        sample.DataVersion,
                        sample.PackageChecksum);
                    RegulationDataModule normModule;
                    if (!bundle.Modules.TryGetValue(
                        RegulationModuleKind.Norm,
                        out normModule))
                    {
                        missing.Add(
                            sample.PackageId + "@" + sample.DataVersion +
                            ": package khong co module Norm.");
                        continue;
                    }
                    catalogs[packageGroup.Key] = NormCatalog.Load(normModule);
                }
                catch (Exception ex) when (
                    ex is System.IO.IOException ||
                    ex is System.IO.InvalidDataException ||
                    ex is ArgumentException ||
                    ex is KeyNotFoundException)
                {
                    missing.Add(
                        sample.PackageId + "@" + sample.DataVersion +
                        ": " + ex.Message);
                }
            }

            EstimateV2WorkItemState[] resolvable = bound
                .Where(item => catalogs.ContainsKey(
                    EstimateV2ResourcePlanBuilder.PackageIdentity(item)))
                .ToArray();

            EstimateV2ResourcePlan plan = EstimateV2ResourcePlanBuilder.Build(
                resolvable,
                item =>
                {
                    string packageIdentity =
                        EstimateV2ResourcePlanBuilder.PackageIdentity(item);
                    return catalogs[packageIdentity].FindRequired(item.NormCode);
                });

            return new WorkbookEstimateV2ResourcePreview(
                plan,
                missing,
                plan.LogicalResources.Select(item => item.Code));
        }

        public static bool TryReadWorkbookUnitPrice(
            Excel.Workbook workbook,
            EstimateV2ResourceRequirement requirement,
            out decimal value)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (requirement == null)
                throw new ArgumentNullException(nameof(requirement));

            foreach (string packageIdentity in requirement.PackageIdentities
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                string name = EstimateV2ExcelNames.ResourcePrice(
                    requirement.Kind,
                    requirement.Code,
                    requirement.Unit,
                    packageIdentity);
                decimal current;
                if (TryReadNamedDecimal(workbook, name, out current) &&
                    current > 0m)
                {
                    value = current;
                    return true;
                }
            }

            value = 0m;
            return false;
        }

        private static bool TryReadNamedDecimal(
            Excel.Workbook workbook,
            string name,
            out decimal value)
        {
            Excel.Names names = null;
            Excel.Name defined = null;
            Excel.Range range = null;
            try
            {
                names = workbook.Names;
                try
                {
                    defined = names.Item(name, Type.Missing, Type.Missing);
                }
                catch (COMException)
                {
                    value = 0m;
                    return false;
                }

                range = defined.RefersToRange;
                object raw = range?.Value2;
                if (raw == null)
                {
                    value = 0m;
                    return false;
                }

                try
                {
                    value = Convert.ToDecimal(raw, CultureInfo.CurrentCulture);
                    return true;
                }
                catch (Exception ex) when (
                    ex is FormatException ||
                    ex is InvalidCastException ||
                    ex is OverflowException)
                {
                    try
                    {
                        value = Convert.ToDecimal(raw, CultureInfo.InvariantCulture);
                        return true;
                    }
                    catch (Exception inner) when (
                        inner is FormatException ||
                        inner is InvalidCastException ||
                        inner is OverflowException)
                    {
                        value = 0m;
                        return false;
                    }
                }
            }
            catch (COMException)
            {
                value = 0m;
                return false;
            }
            finally
            {
                Release(range);
                Release(defined);
                Release(names);
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

    }
}
