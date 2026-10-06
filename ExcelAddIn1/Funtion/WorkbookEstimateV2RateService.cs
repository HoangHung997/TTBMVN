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
    public sealed class WorkbookEstimateV2RateItemPreview
    {
        internal WorkbookEstimateV2RateItemPreview(
            EstimateV2RateItem rate,
            int missingPriceCount,
            int formulaLinkCount,
            bool isGenerated,
            IEnumerable<string> warnings)
        {
            Rate = rate ?? throw new ArgumentNullException(nameof(rate));
            MissingPriceCount = missingPriceCount;
            FormulaLinkCount = formulaLinkCount;
            IsGenerated = isGenerated;
            Warnings = new ReadOnlyCollection<string>(
                (warnings ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public EstimateV2RateItem Rate { get; }
        public int MissingPriceCount { get; }
        public int FormulaLinkCount { get; }
        public bool IsGenerated { get; }
        public IReadOnlyList<string> Warnings { get; }
        public bool HasMissingPrice => MissingPriceCount > 0;
        public bool HasWarning => Warnings.Count > 0;
    }

    public sealed class WorkbookEstimateV2RatePreview
    {
        internal WorkbookEstimateV2RatePreview(
            EstimateV2RateEnvironment environment,
            IEnumerable<WorkbookEstimateV2RateItemPreview> items,
            IEnumerable<string> missingPackageBindings)
        {
            Environment = environment;
            Items = new ReadOnlyCollection<WorkbookEstimateV2RateItemPreview>(
                (items ?? Enumerable.Empty<WorkbookEstimateV2RateItemPreview>())
                    .OrderBy(item => item.Rate.NormCode, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(item => item.Rate.VariantCode, StringComparer.OrdinalIgnoreCase)
                    .ToList());
            MissingPackageBindings = new ReadOnlyCollection<string>(
                (missingPackageBindings ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public EstimateV2RateEnvironment Environment { get; }
        public IReadOnlyList<WorkbookEstimateV2RateItemPreview> Items { get; }
        public IReadOnlyList<string> MissingPackageBindings { get; }

        public int NeededCount => Items.Count;
        public int GeneratedCount => Items.Count(item => item.IsGenerated);
        public int MissingRateCount => Items.Count(item => item.HasMissingPrice);
        public int WarningRateCount =>
            Items.Count(item => item.HasWarning) + MissingPackageBindings.Count;
        public int FormulaLinkCount => Items.Sum(item => item.FormulaLinkCount);

        public bool CanGenerate =>
            Items.Count > 0 &&
            MissingPackageBindings.Count == 0;
    }

    public static class WorkbookEstimateV2RateService
    {
        public static WorkbookEstimateV2RatePreview BuildPreview(
            Excel.Workbook workbook,
            EstimateV2RateEnvironment environment)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            EstimateV2State state;
            if (!WorkbookEstimateV2StateService.TryLoad(workbook, out state))
            {
                return new WorkbookEstimateV2RatePreview(
                    environment,
                    Enumerable.Empty<WorkbookEstimateV2RateItemPreview>(),
                    Enumerable.Empty<string>());
            }

            EstimateV2WorkItemState[] bound = state.WorkItems
                .Where(item =>
                    item != null &&
                    !item.IsOrphaned &&
                    item.HasNormBinding &&
                    SafeEnvironment(item.NormCode) == environment)
                .ToArray();
            if (bound.Length == 0)
            {
                return new WorkbookEstimateV2RatePreview(
                    environment,
                    Enumerable.Empty<WorkbookEstimateV2RateItemPreview>(),
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
                        ": binding chua pin day du package. Hay gan lai dinh muc.");
                    continue;
                }

                try
                {
                    RegulationPackageBundle bundle = store.LoadBundleRequired(
                        sample.PackageId,
                        sample.DataVersion,
                        sample.PackageChecksum);
                    RegulationDataModule module;
                    if (!bundle.Modules.TryGetValue(
                        RegulationModuleKind.Norm,
                        out module))
                    {
                        missing.Add(
                            sample.PackageId + "@" + sample.DataVersion +
                            ": package khong co module Norm.");
                        continue;
                    }
                    catalogs[packageGroup.Key] = NormCatalog.Load(module);
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

            EstimateV2RatePlan plan = EstimateV2RatePlanBuilder.Build(
                resolvable,
                item =>
                {
                    string packageIdentity =
                        EstimateV2ResourcePlanBuilder.PackageIdentity(item);
                    return catalogs[packageIdentity].FindRequired(item.NormCode);
                });

            var previews = new List<WorkbookEstimateV2RateItemPreview>();
            foreach (EstimateV2RateItem rate in plan.ForEnvironment(environment))
            {
                int missingPrices = 0;
                int links = 0;
                var warnings = new List<string>();

                foreach (EstimateV2RateResource resource in rate.Resources)
                {
                    if (resource.IsPercentage)
                        continue;
                    links++;

                    bool hasPrice = resource.PriceCandidates.Any(candidate =>
                    {
                        decimal value;
                        return TryReadResourcePrice(
                            workbook,
                            resource.Kind,
                            candidate,
                            resource.Unit,
                            rate.PackageIdentity,
                            out value) &&
                            value > 0m;
                    });
                    if (!hasPrice)
                        missingPrices++;

                    if (resource.PriceCandidates.Count > 1)
                    {
                        warnings.Add(
                            resource.ResourceCode +
                            ": co nhieu phuong an gia; hien tai se dung gia kha dung dau tien.");
                    }
                }

                if (rate.RequiresConditionReview)
                {
                    warnings.Add(
                        "Dinh muc co he so/dieu kien; V2-301 dang sinh bo hao phi co so theo variant da gan.");
                }

                previews.Add(new WorkbookEstimateV2RateItemPreview(
                    rate,
                    missingPrices,
                    links,
                    IsGenerated(workbook, rate.RateId),
                    warnings));
            }

            return new WorkbookEstimateV2RatePreview(
                environment,
                previews,
                missing);
        }

        public static bool TryReadResourcePrice(
            Excel.Workbook workbook,
            NormResourceKind kind,
            string resourceCode,
            string unit,
            string packageIdentity,
            out decimal value)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            string name = EstimateV2ExcelNames.ResourcePrice(
                kind,
                resourceCode,
                unit,
                packageIdentity);
            return TryReadNamedDecimal(workbook, name, out value);
        }

        public static bool IsGenerated(
            Excel.Workbook workbook,
            string rateId)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            return HasWorkbookName(
                workbook,
                EstimateV2ExcelNames.RateComponent(rateId, "VL")) &&
                HasWorkbookName(
                    workbook,
                    EstimateV2ExcelNames.RateComponent(rateId, "NC")) &&
                HasWorkbookName(
                    workbook,
                    EstimateV2ExcelNames.RateComponent(rateId, "M"));
        }

        private static EstimateV2RateEnvironment SafeEnvironment(string normCode)
        {
            try
            {
                return EstimateV2RatePlanBuilder.ClassifyEnvironment(normCode);
            }
            catch (FormatException)
            {
                return EstimateV2RateEnvironment.Land;
            }
        }

        private static bool HasWorkbookName(
            Excel.Workbook workbook,
            string name)
        {
            Excel.Names names = null;
            Excel.Name defined = null;
            try
            {
                names = workbook.Names;
                try
                {
                    defined = names.Item(name, Type.Missing, Type.Missing);
                    return defined != null;
                }
                catch (COMException)
                {
                    return false;
                }
            }
            finally
            {
                Release(defined);
                Release(names);
            }
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

                try
                {
                    range = defined.RefersToRange;
                }
                catch (COMException)
                {
                    value = 0m;
                    return false;
                }

                object raw = range?.Value2;
                if (raw == null)
                {
                    value = 0m;
                    return true;
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
