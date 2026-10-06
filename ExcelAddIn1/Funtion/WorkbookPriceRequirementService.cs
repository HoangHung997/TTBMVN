using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class WorkbookPriceRequirementService
    {
        public static IReadOnlyList<PriceRequirement> LoadPinnedNormRequirements(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            RegulationPackageBootstrapService.LoadAvailablePackages();
            ProjectProfile project = WorkbookProjectProfileService.LoadRequired(workbook);
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            RegulationPackageBundle bundle = store.LoadBundleRequired(
                project.RegulationPackageId,
                project.RegulationPackageVersion,
                project.RegulationPackageChecksum);
            RegulationDataModule module;
            if (!bundle.Modules.TryGetValue(RegulationModuleKind.Norm, out module))
                return new ReadOnlyCollection<PriceRequirement>(new List<PriceRequirement>());

            NormCatalog catalog;
            try
            {
                catalog = NormCatalog.Load(module);
            }
            catch (FormatException)
            {
                return new ReadOnlyCollection<PriceRequirement>(new List<PriceRequirement>());
            }

            var result = new Dictionary<string, PriceRequirement>(StringComparer.OrdinalIgnoreCase);
            foreach (NormResourceRate rate in catalog.Definitions.SelectMany(item => item.Rates))
            {
                PriceResourceKind kind;
                switch (rate.Kind)
                {
                    case NormResourceKind.Material:
                        if (string.Equals(rate.Unit, "percent", StringComparison.OrdinalIgnoreCase))
                            continue;
                        kind = PriceResourceKind.Material;
                        break;
                    case NormResourceKind.Labor:
                        kind = PriceResourceKind.Labor;
                        break;
                    case NormResourceKind.Machine:
                        kind = PriceResourceKind.MachineShift;
                        break;
                    default:
                        continue;
                }
                string identity = ((int)kind) + "|" + rate.ResourceCode + "|" + rate.Unit;
                if (!result.ContainsKey(identity))
                    result.Add(identity, new PriceRequirement(rate.ResourceCode, kind, rate.Unit));
            }
            return new ReadOnlyCollection<PriceRequirement>(result.Values
                .OrderBy(item => item.Kind)
                .ThenBy(item => item.Code, StringComparer.OrdinalIgnoreCase)
                .ToList());
        }
    }
}
