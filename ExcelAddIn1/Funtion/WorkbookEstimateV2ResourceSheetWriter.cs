using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookEstimateV2ResourceWriteResult
    {
        internal WorkbookEstimateV2ResourceWriteResult(
            string worksheetName,
            int materialCount,
            int laborCount,
            int machineCount,
            int formulaCount,
            int inputCount,
            int missingInputCount)
        {
            WorksheetName = worksheetName ?? string.Empty;
            MaterialCount = materialCount;
            LaborCount = laborCount;
            MachineCount = machineCount;
            FormulaCount = formulaCount;
            InputCount = inputCount;
            MissingInputCount = missingInputCount;
        }

        public string WorksheetName { get; }
        public int MaterialCount { get; }
        public int LaborCount { get; }
        public int MachineCount { get; }
        public int FormulaCount { get; }
        public int InputCount { get; }
        public int MissingInputCount { get; }
    }

    public static class WorkbookEstimateV2ResourceSheetWriter
    {
        private const string SheetName = "VL-NC-M";
        private const string GeneratedProperty = "TTBMVN.EstimateV2.ResourcePrices";
        private const string GeneratedVersion = "1";
        private const int VisibleLastColumn = 6; // A:F
        private const int MetadataColumnCount = 11;

        private static readonly string[] MetadataHeaders =
        {
            "__TTB_ROWTYPE",
            "__TTB_RESOURCE",
            "__TTB_KIND",
            "__TTB_PACKAGE",
            "__TTB_INPUTKEY",
            "__TTB_PARAM1",
            "__TTB_PARAM2",
            "__TTB_PARAM3",
            "__TTB_PARAM4",
            "__TTB_PARAM5",
            "__TTB_PARAM6"
        };

        public static WorkbookEstimateV2ResourceWriteResult Apply(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            WorkbookEstimateV2ResourcePreview preview =
                WorkbookEstimateV2ResourceService.BuildPreview(workbook);
            if (preview.Plan.BoundWorkItemCount == 0)
                throw new InvalidOperationException(
                    "Chua co cong tac duoc gan dinh muc de sinh VL-NC-M.");
            if (preview.MissingPackageBindings.Count > 0)
                throw new InvalidOperationException(
                    "Chua resolve du package cua binding: " +
                    string.Join(" ", preview.MissingPackageBindings));
            if (preview.UnresolvedLogicalResources.Count > 0)
                throw new InvalidOperationException(
                    "Con tai nguyen logic chua rang buoc: " +
                    string.Join(", ", preview.UnresolvedLogicalResources));

            EstimateV2State state;
            if (!WorkbookEstimateV2StateService.TryLoad(workbook, out state))
                throw new InvalidOperationException("Khong doc duoc Estimate V2 state.");

            EstimateV2WorkItemState[] bound = state.WorkItems
                .Where(item => item != null && !item.IsOrphaned && item.HasNormBinding)
                .ToArray();
            string[] packageIdentities = bound
                .Select(EstimateV2ResourcePlanBuilder.PackageIdentity)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (packageIdentities.Length != 1)
            {
                throw new InvalidOperationException(
                    "V2-201 hien tai chi sinh mot bo VL-NC-M cho mot package phap ly trong workbook. " +
                    "Phat hien " + packageIdentities.Length + " package.");
            }

            PriceProfilePortfolio portfolio;
            WorkbookPriceProfilePortfolioService.TryLoad(workbook, out portfolio);
            PriceProfile profile = ResolveProfile(portfolio, preview.Plan);

            EstimateV2WorkItemState sample = bound[0];
            RegulationPackageBootstrapService.LoadAvailablePackages();
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            RegulationPackageBundle bundle = store.LoadBundleRequired(
                sample.PackageId,
                sample.DataVersion,
                sample.PackageChecksum);

            MachineRateCatalog machineCatalog = null;
            if (preview.Plan.Machines.Count > 0)
            {
                if (profile == null)
                {
                    throw new InvalidOperationException(
                        "Can xac dinh doi tuong luong (HLNS/KHLNS) truoc khi sinh gia ca may. " +
                        "Hay tao/chon mot ho so gia hoac chon doi tuong luong trong Thiet lap.");
                }

                RegulationDataModule machineModule;
                if (!bundle.Modules.TryGetValue(
                    RegulationModuleKind.MachineRate,
                    out machineModule))
                {
                    throw new InvalidOperationException(
                        "Package dang dung khong co module MachineRate.");
                }
                machineCatalog = MachineRateCatalog.Load(
                    machineModule,
                    profile.LaborAudience);
            }

            Excel.Worksheet sheet = null;
            bool created = false;
            try
            {
                sheet = ResolveOrCreateSheet(workbook, out created);
                WorksheetRoleService.SetRole(sheet, WorksheetRole.ResourcePrices);
                SetWorksheetProperty(sheet, GeneratedProperty, GeneratedVersion);

                var preserved = CapturePreservedInputs(
                    workbook,
                    sheet,
                    preview.Plan,
                    profile);

                ResourceSheetBuild build = BuildSheet(
                    workbook,
                    sheet,
                    preview.Plan,
                    packageIdentities[0],
                    profile,
                    machineCatalog,
                    preserved);

                using (new ExcelWriteContext(workbook.Application))
                using (var transaction = new ExcelBatchWriteTransaction())
                {
                    UnmergeVisibleArea(sheet, Math.Max(build.LastRow, ExistingLastRow(sheet)));
                    ClearManagedArea(sheet, build.LastRow, build.MetadataStartColumn, transaction);
                    WriteMatrix(
                        sheet,
                        build,
                        transaction);
                    ApplyNames(
                        workbook,
                        sheet,
                        build);
                    workbook.Application.Calculate();
                    transaction.Commit();
                }

                FormatSheet(
                    sheet,
                    build);
                HideTechnicalColumns(
                    sheet,
                    build.MetadataStartColumn,
                    build.MetadataStartColumn + MetadataColumnCount - 1);
                sheet.PageSetup.PrintArea = "$A$1:$F$" +
                    build.LastRow.ToString(CultureInfo.InvariantCulture);

                return new WorkbookEstimateV2ResourceWriteResult(
                    sheet.Name,
                    preview.Plan.Materials.Count,
                    build.LaborCount,
                    preview.Plan.Machines.Count,
                    build.FormulaCount,
                    build.InputCount,
                    build.MissingInputCount);
            }
            catch
            {
                if (created && sheet != null)
                DeleteSheetQuietly(workbook, sheet);
                throw;
            }
            finally
            {
                Release(sheet);
            }
        }

        private static PriceProfile ResolveProfile(
            PriceProfilePortfolio portfolio,
            EstimateV2ResourcePlan plan)
        {
            if (portfolio == null || portfolio.Profiles.Count == 0)
                return null;
            if (portfolio.Profiles.Count == 1)
                return portfolio.Profiles[0];

            // Khong tu y doan khi workbook co dong thoi HLNS va KHLNS.
            throw new InvalidOperationException(
                "Workbook co nhieu ho so gia HLNS/KHLNS. " +
                "Hay chon doi tuong luong trong Thiet lap truoc khi sinh VL-NC-M.");
        }

        private static ResourceSheetBuild BuildSheet(
            Excel.Workbook workbook,
            Excel.Worksheet sheet,
            EstimateV2ResourcePlan plan,
            string packageIdentity,
            PriceProfile profile,
            MachineRateCatalog machineCatalog,
            PreservedInputs preserved)
        {
            int existingLastColumn = ExistingLastColumn(sheet);
            int metadataStart = Math.Max(8, existingLastColumn + 1);
            if (metadataStart + MetadataColumnCount - 1 > 16384)
                throw new InvalidOperationException("Khong con cot trong de luu metadata VL-NC-M.");

            var rows = new List<ResourceSheetRow>();
            int formulaCount = 0;
            int inputCount = 0;
            int missingInputCount = 0;

            var machineDefinitions = new Dictionary<string, MachineRateDefinition>(
                StringComparer.OrdinalIgnoreCase);
            var laborRequirements = plan.Labor.ToList();
            if (machineCatalog != null)
            {
                foreach (EstimateV2ResourceRequirement machine in plan.Machines)
                {
                    MachineRateDefinition definition = ResolveMachine(
                        machineCatalog,
                        machine.Code);
                    machineDefinitions[machine.Code] = definition;
                    foreach (MachineOperatorRequirement op in definition.Operators)
                    {
                        string normalized = NormalizeLaborCode(op.LaborCode);
                        if (laborRequirements.Any(item => string.Equals(
                            NormalizeLaborCode(item.Code),
                            normalized,
                            StringComparison.OrdinalIgnoreCase)))
                            continue;

                        laborRequirements.Add(new EstimateV2ResourceRequirement(
                            NormResourceKind.Labor,
                            normalized,
                            "worker-day",
                            true,
                            false,
                            1,
                            Enumerable.Empty<string>(),
                            new[] { packageIdentity }));
                    }
                }
            }

            rows.Add(ResourceSheetRow.Title("CÁC PHỤ LỤC"));
            rows.Add(ResourceSheetRow.Title("GIÁ NHÂN CÔNG, CA MÁY VÀ VẬT LIỆU"));
            rows.Add(ResourceSheetRow.Section("I. GIÁ NHÂN CÔNG"));
            rows.Add(ResourceSheetRow.Header(new[]
            {
                "TT",
                "Nội dung đơn giá",
                "Hệ số",
                "Lương tối thiểu (đồng)",
                "Ngày công",
                "Thành tiền (đ)"
            }));

            var priceNames = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            var namedCells = new List<NamedCell>();

            int laborIndex = 0;
            foreach (EstimateV2ResourceRequirement labor in laborRequirements
                .OrderBy(item => item.Code, StringComparer.OrdinalIgnoreCase))
            {
                laborIndex++;
                int headerRow = rows.Count + 1;
                string displayName = ResourceDisplayName(profile, labor.Code);
                rows.Add(ResourceSheetRow.LaborHeader(
                    Roman(laborIndex),
                    displayName,
                    labor.Code,
                    packageIdentity));

                if (IsQncnLabor(labor.Code))
                {
                    LaborInputSet inputs = preserved.GetLabor(
                        labor.Code,
                        profile);
                    int baseRow = rows.Count + 1;
                    rows.Add(ResourceSheetRow.LaborInput(
                        "1",
                        "Lương cơ bản " + displayName,
                        inputs.Coefficient,
                        inputs.BaseSalary,
                        inputs.WorkDays,
                        "=IFERROR(D" + baseRow + "*C" + baseRow + "/E" + baseRow + ",0)",
                        labor.Code,
                        packageIdentity,
                        "BASE"));
                    formulaCount++;
                    inputCount += 3;
                    missingInputCount += Missing(inputs.Coefficient) +
                        Missing(inputs.BaseSalary) +
                        Missing(inputs.WorkDays);

                    int dangerRow = rows.Count + 1;
                    rows.Add(ResourceSheetRow.LaborInput(
                        "2",
                        "Phụ cấp khó khăn nguy hiểm",
                        inputs.DangerAllowance,
                        "=" + ExcelColumnAddress.ToLetters(4) + baseRow,
                        inputs.WorkDays,
                        "=IFERROR(D" + dangerRow + "*C" + dangerRow + "/E" + dangerRow + ",0)",
                        labor.Code,
                        packageIdentity,
                        "DANGER"));
                    formulaCount += 2;
                    inputCount += 1;
                    missingInputCount += Missing(inputs.DangerAllowance);

                    int mobileRow = rows.Count + 1;
                    rows.Add(ResourceSheetRow.LaborInput(
                        "3",
                        "Phụ cấp lưu động",
                        inputs.MobileAllowance,
                        "=" + ExcelColumnAddress.ToLetters(4) + baseRow,
                        inputs.WorkDays,
                        "=IFERROR(D" + mobileRow + "*C" + mobileRow + "/E" + mobileRow + ",0)",
                        labor.Code,
                        packageIdentity,
                        "MOBILE"));
                    formulaCount += 2;
                    inputCount += 1;
                    missingInputCount += Missing(inputs.MobileAllowance);

                    rows[headerRow - 1].Cells[5] =
                        "=SUM(F" + baseRow + ":F" + mobileRow + ")";
                    formulaCount++;

                    string priceName = EstimateV2ExcelNames.ResourcePrice(
                        NormResourceKind.Labor,
                        labor.Code,
                        labor.Unit,
                        packageIdentity);
                    namedCells.Add(new NamedCell(priceName, headerRow, 6));
                    priceNames[ResourceKey(NormResourceKind.Labor, labor.Code, labor.Unit)] =
                        priceName;

                    namedCells.Add(new NamedCell(
                        EstimateV2ExcelNames.LaborInput(labor.Code, "COEFFICIENT"),
                        baseRow,
                        3));
                    namedCells.Add(new NamedCell(
                        EstimateV2ExcelNames.LaborInput(labor.Code, "BASESALARY"),
                        baseRow,
                        4));
                    namedCells.Add(new NamedCell(
                        EstimateV2ExcelNames.LaborInput(labor.Code, "WORKDAYS"),
                        baseRow,
                        5));
                    namedCells.Add(new NamedCell(
                        EstimateV2ExcelNames.LaborInput(labor.Code, "DANGER"),
                        dangerRow,
                        3));
                    namedCells.Add(new NamedCell(
                        EstimateV2ExcelNames.LaborInput(labor.Code, "MOBILE"),
                        mobileRow,
                        3));
                }
                else
                {
                    decimal? dailyInput = preserved.GetDirectLaborRate(
                        labor.Code,
                        profile);
                    int inputRow = rows.Count + 1;
                    rows.Add(ResourceSheetRow.GenericLaborInput(
                        "1",
                        "Đơn giá ngày công " + displayName,
                        dailyInput,
                        "=IFERROR(D" + inputRow + ",0)",
                        labor.Code,
                        packageIdentity));
                    inputCount++;
                    missingInputCount += Missing(dailyInput);
                    formulaCount++;

                    rows[headerRow - 1].Cells[5] = "=F" + inputRow;
                    formulaCount++;

                    string priceName = EstimateV2ExcelNames.ResourcePrice(
                        NormResourceKind.Labor,
                        labor.Code,
                        labor.Unit,
                        packageIdentity);
                    namedCells.Add(new NamedCell(priceName, headerRow, 6));
                    priceNames[ResourceKey(NormResourceKind.Labor, labor.Code, labor.Unit)] =
                        priceName;
                }
            }

            rows.Add(ResourceSheetRow.Section("II. GIÁ CA MÁY"));
            rows.Add(ResourceSheetRow.Header(new[]
            {
                "TT",
                "Nội dung chi phí",
                "ĐVT",
                "Khối lượng",
                "Đ.Giá",
                "Thành tiền (đ)"
            }));

            int machineIndex = 0;
            var fuelInputs = new Dictionary<string, FuelInputDefinition>(
                StringComparer.OrdinalIgnoreCase);
            foreach (EstimateV2ResourceRequirement machine in plan.Machines)
            {
                machineIndex++;
                if (machineCatalog == null)
                    throw new InvalidOperationException("Thieu MachineRateCatalog.");

                MachineRateDefinition definition = machineDefinitions[machine.Code];
                int headerRow = rows.Count + 1;
                rows.Add(ResourceSheetRow.MachineHeader(
                    Roman(machineIndex),
                    definition.Title,
                    machine.Code,
                    packageIdentity));

                int firstDetail = rows.Count + 1;
                if (!definition.Fuel.IsNone)
                {
                    FuelInputDefinition fuel;
                    if (!fuelInputs.TryGetValue(definition.Fuel.PriceCode, out fuel))
                    {
                        fuel = new FuelInputDefinition(
                            definition.Fuel.PriceCode,
                            definition.Fuel.Unit,
                            preserved.GetFuelPrice(
                                definition.Fuel.PriceCode,
                                profile));
                        fuelInputs.Add(definition.Fuel.PriceCode, fuel);
                    }

                    int row = rows.Count + 1;
                    string fuelName = EstimateV2ExcelNames.FuelInput(
                        definition.Fuel.PriceCode);
                    rows.Add(ResourceSheetRow.MachineDetail(
                        (row - firstDetail + 1).ToString(CultureInfo.InvariantCulture),
                        FuelDisplayName(profile, definition.Fuel.PriceCode),
                        definition.Fuel.Unit,
                        definition.Fuel.Quantity,
                        "=" + fuelName,
                        "=IFERROR(D" + row + "*E" + row + "*" +
                            HiddenParamReference(metadataStart, row, 8) + ",0)",
                        machine.Code,
                        packageIdentity,
                        "FUEL",
                        definition.Fuel.Kind.ToString(),
                        definition.Fuel.PriceCode,
                        definition.Fuel.GetAuxiliaryFactorForFormula()));
                    formulaCount += 2;
                }

                decimal corrosionFactor = IsCorrosiveMachineContext(machine)
                    ? MachineRateCalculator.CorrosiveEnvironmentFactor
                    : 1m;

                int repairRow = rows.Count + 1;
                rows.Add(ResourceSheetRow.MachineDetail(
                    (repairRow - firstDetail + 1).ToString(CultureInfo.InvariantCulture),
                    "Sửa chữa, bảo dưỡng",
                    "Ca",
                    1m,
                    "=IFERROR(" + HiddenParamReference(metadataStart, repairRow, 6) +
                        "*(" + HiddenParamReference(metadataStart, repairRow, 7) +
                        "*" + HiddenParamReference(metadataStart, repairRow, 9) +
                        ")/100/" + HiddenParamReference(metadataStart, repairRow, 8) + ",0)",
                    "=D" + repairRow + "*E" + repairRow,
                    machine.Code,
                    packageIdentity,
                    "REPAIR",
                    definition.ReferencePriceVnd,
                    definition.RepairPercent,
                    definition.AnnualShifts,
                    corrosionFactor));
                formulaCount += 2;

                int depreciationRow = rows.Count + 1;
                rows.Add(ResourceSheetRow.MachineDetail(
                    (depreciationRow - firstDetail + 1).ToString(CultureInfo.InvariantCulture),
                    "Khấu hao cơ bản",
                    "Ca",
                    1m,
                    "=IFERROR((" + HiddenParamReference(metadataStart, depreciationRow, 6) +
                        "-(" + HiddenParamReference(metadataStart, depreciationRow, 6) +
                        "*" + HiddenParamReference(metadataStart, depreciationRow, 7) +
                        "/100))*(" + HiddenParamReference(metadataStart, depreciationRow, 8) +
                        "*" + HiddenParamReference(metadataStart, depreciationRow, 10) +
                        ")/100/" + HiddenParamReference(metadataStart, depreciationRow, 9) + ",0)",
                    "=D" + depreciationRow + "*E" + depreciationRow,
                    machine.Code,
                    packageIdentity,
                    "DEPRECIATION",
                    definition.ReferencePriceVnd,
                    definition.ReferenceRecoverableValuePercent,
                    definition.DepreciationPercent,
                    definition.AnnualShifts,
                    corrosionFactor));
                formulaCount += 2;

                int otherRow = rows.Count + 1;
                rows.Add(ResourceSheetRow.MachineDetail(
                    (otherRow - firstDetail + 1).ToString(CultureInfo.InvariantCulture),
                    "Chi phí khác",
                    "Ca",
                    1m,
                    "=IFERROR(" + HiddenParamReference(metadataStart, otherRow, 6) +
                        "*" + HiddenParamReference(metadataStart, otherRow, 7) +
                        "/100/" + HiddenParamReference(metadataStart, otherRow, 8) + ",0)",
                    "=D" + otherRow + "*E" + otherRow,
                    machine.Code,
                    packageIdentity,
                    "OTHER",
                    definition.ReferencePriceVnd,
                    definition.OtherPercent,
                    definition.AnnualShifts));
                formulaCount += 2;

                foreach (MachineOperatorRequirement op in definition.Operators)
                {
                    int row = rows.Count + 1;
                    EstimateV2ResourceRequirement labor = FindLaborForOperator(
                        laborRequirements,
                        op.LaborCode);
                    string laborPriceName = EstimateV2ExcelNames.ResourcePrice(
                        NormResourceKind.Labor,
                        labor.Code,
                        labor.Unit,
                        packageIdentity);
                    rows.Add(ResourceSheetRow.MachineDetail(
                        (row - firstDetail + 1).ToString(CultureInfo.InvariantCulture),
                        "Nhân công " + ResourceDisplayName(profile, labor.Code),
                        labor.Unit.Length > 0 ? labor.Unit : "Công",
                        op.Quantity,
                        "=" + laborPriceName,
                        "=D" + row + "*E" + row,
                        machine.Code,
                        packageIdentity,
                        "OPERATOR",
                        op.LaborCode));
                    formulaCount += 2;
                }

                int lastDetail = rows.Count;
                rows[headerRow - 1].Cells[5] =
                    lastDetail >= firstDetail
                        ? "=SUM(F" + firstDetail + ":F" + lastDetail + ")"
                        : "=0";
                formulaCount++;

                string machinePriceName = EstimateV2ExcelNames.ResourcePrice(
                    NormResourceKind.Machine,
                    machine.Code,
                    machine.Unit,
                    packageIdentity);
                namedCells.Add(new NamedCell(machinePriceName, headerRow, 6));
                priceNames[ResourceKey(
                    NormResourceKind.Machine,
                    machine.Code,
                    machine.Unit)] = machinePriceName;
            }

            if (fuelInputs.Count > 0)
            {
                rows.Add(ResourceSheetRow.Note("Ghi chú:"));
                rows.Add(ResourceSheetRow.Note(
                    "Giá nhiên liệu/năng lượng được lưu ở vùng kỹ thuật ẩn và được công thức giá ca máy liên kết trực tiếp."));
            }

            rows.Add(ResourceSheetRow.Section("III. GIÁ VẬT LIỆU"));
            rows.Add(ResourceSheetRow.Header(new[]
            {
                "TT",
                "Nội dung",
                "",
                "Ghi chú",
                "ĐVT",
                "Đơn giá (đ)"
            }));

            int materialIndex = 0;
            foreach (EstimateV2ResourceRequirement material in plan.Materials
                .Where(item => item.RequiresUnitPrice))
            {
                materialIndex++;
                decimal? price = preserved.GetMaterialPrice(material.Code, profile);
                int row = rows.Count + 1;
                rows.Add(ResourceSheetRow.MaterialInput(
                    materialIndex,
                    ResourceDisplayName(profile, material.Code),
                    material.Unit,
                    price,
                    material.Code,
                    packageIdentity));
                inputCount++;
                missingInputCount += Missing(price);

                string materialPriceName = EstimateV2ExcelNames.ResourcePrice(
                    NormResourceKind.Material,
                    material.Code,
                    material.Unit,
                    packageIdentity);
                namedCells.Add(new NamedCell(materialPriceName, row, 6));
                priceNames[ResourceKey(
                    NormResourceKind.Material,
                    material.Code,
                    material.Unit)] = materialPriceName;
            }

            int lastRow = rows.Count;
            int hiddenInputStartRow = 2;
            foreach (FuelInputDefinition fuel in fuelInputs.Values
                .OrderBy(item => item.Code, StringComparer.OrdinalIgnoreCase))
            {
                fuel.HiddenRow = hiddenInputStartRow++;
                string name = EstimateV2ExcelNames.FuelInput(fuel.Code);
                namedCells.Add(new NamedCell(
                    name,
                    fuel.HiddenRow,
                    metadataStart + 5));
                inputCount++;
                missingInputCount += Missing(fuel.Price);
            }

            return new ResourceSheetBuild(
                rows,
                lastRow,
                metadataStart,
                namedCells,
                fuelInputs.Values,
                formulaCount,
                inputCount,
                missingInputCount,
                laborRequirements.Count);
        }

        private static PreservedInputs CapturePreservedInputs(
            Excel.Workbook workbook,
            Excel.Worksheet sheet,
            EstimateV2ResourcePlan plan,
            PriceProfile profile)
        {
            var result = new PreservedInputs();

            string packageIdentity = plan.Resources
                .SelectMany(item => item.PackageIdentities)
                .FirstOrDefault() ?? string.Empty;

            foreach (EstimateV2ResourceRequirement material in plan.Materials)
            {
                decimal? value = ReadNameDecimal(
                    workbook,
                    EstimateV2ExcelNames.ResourcePrice(
                        NormResourceKind.Material,
                        material.Code,
                        material.Unit,
                        packageIdentity));
                if (value.HasValue)
                    result.MaterialPrices[material.Code] = value.Value;
            }

            foreach (EstimateV2ResourceRequirement labor in plan.Labor)
            {
                LaborInputSet inputs = new LaborInputSet();
                inputs.Coefficient = ReadNameDecimal(
                    workbook,
                    EstimateV2ExcelNames.LaborInput(labor.Code, "COEFFICIENT"));
                inputs.BaseSalary = ReadNameDecimal(
                    workbook,
                    EstimateV2ExcelNames.LaborInput(labor.Code, "BASESALARY"));
                inputs.WorkDays = ReadNameDecimal(
                    workbook,
                    EstimateV2ExcelNames.LaborInput(labor.Code, "WORKDAYS"));
                inputs.DangerAllowance = ReadNameDecimal(
                    workbook,
                    EstimateV2ExcelNames.LaborInput(labor.Code, "DANGER"));
                inputs.MobileAllowance = ReadNameDecimal(
                    workbook,
                    EstimateV2ExcelNames.LaborInput(labor.Code, "MOBILE"));
                decimal? direct = ReadNameDecimal(
                    workbook,
                    EstimateV2ExcelNames.ResourcePrice(
                        NormResourceKind.Labor,
                        labor.Code,
                        labor.Unit,
                        packageIdentity));
                if (direct.HasValue)
                    result.DirectLaborRates[labor.Code] = direct.Value;
                if (inputs.HasAny)
                    result.LaborInputs[labor.Code] = inputs;
            }

            CaptureLegacyLaborInputs(sheet, result);
            CaptureLegacyMaterialPrices(sheet, result, profile);

            if (profile != null)
            {
                foreach (PriceProfileEntry entry in profile.Entries
                    .Where(item => item.Kind == PriceResourceKind.FuelEnergy))
                {
                    PriceProfilePrice price = profile.FindRequired(entry.Code);
                    result.FuelPrices[entry.Code] = price.AppliedUnitPriceVnd;
                }
            }

            return result;
        }

        private static void CaptureLegacyLaborInputs(
            Excel.Worksheet sheet,
            PreservedInputs result)
        {
            int lastRow = Math.Min(200, ExistingLastRow(sheet));
            for (int row = 1; row <= lastRow; row++)
            {
                string description = ReadText(sheet, row, 2);
                string laborCode = LaborCodeFromDescription(description);
                if (laborCode.Length == 0)
                    continue;

                LaborInputSet inputs;
                if (!result.LaborInputs.TryGetValue(laborCode, out inputs))
                {
                    inputs = new LaborInputSet();
                    result.LaborInputs[laborCode] = inputs;
                }

                if (description.IndexOf(
                    "Lương cơ bản",
                    StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    inputs.Coefficient = inputs.Coefficient ?? ReadDecimal(sheet, row, 3);
                    inputs.BaseSalary = inputs.BaseSalary ?? ReadDecimal(sheet, row, 4);
                    inputs.WorkDays = inputs.WorkDays ?? ReadDecimal(sheet, row, 5);
                }
                else if (description.IndexOf(
                    "khó khăn",
                    StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    inputs.DangerAllowance = inputs.DangerAllowance ?? ReadDecimal(sheet, row, 3);
                }
                else if (description.IndexOf(
                    "lưu động",
                    StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    inputs.MobileAllowance = inputs.MobileAllowance ?? ReadDecimal(sheet, row, 3);
                }
            }
        }

        private static void CaptureLegacyMaterialPrices(
            Excel.Worksheet sheet,
            PreservedInputs result,
            PriceProfile profile)
        {
            if (profile == null)
                return;

            int lastRow = ExistingLastRow(sheet);
            var byName = profile.Entries
                .Where(item => item.Kind == PriceResourceKind.Material)
                .Where(item => !string.IsNullOrWhiteSpace(item.DisplayName))
                .GroupBy(item => NormalizeText(item.DisplayName), StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() == 1)
                .ToDictionary(
                    group => group.Key,
                    group => group.First().Code,
                    StringComparer.OrdinalIgnoreCase);

            for (int row = 1; row <= lastRow; row++)
            {
                string name = NormalizeText(ReadText(sheet, row, 2));
                string code;
                if (name.Length == 0 || !byName.TryGetValue(name, out code))
                    continue;
                decimal? price = ReadDecimal(sheet, row, 6);
                if (price.HasValue && price.Value >= 0m)
                    result.MaterialPrices[code] = price.Value;
            }
        }

        private static MachineRateDefinition ResolveMachine(
            MachineRateCatalog catalog,
            string resourceCode)
        {
            string code = (resourceCode ?? string.Empty).Trim();
            if (code.StartsWith("M010.", StringComparison.OrdinalIgnoreCase))
                return catalog.FindRequiredByKey("MACHINE-" + code.ToUpperInvariant());
            return catalog.FindRequiredByCode(code);
        }

        private static EstimateV2ResourceRequirement FindLaborForOperator(
            IEnumerable<EstimateV2ResourceRequirement> laborRequirements,
            string operatorCode)
        {
            string normalized = NormalizeLaborCode(operatorCode);
            EstimateV2ResourceRequirement match =
                (laborRequirements ?? Enumerable.Empty<EstimateV2ResourceRequirement>())
                .FirstOrDefault(item => string.Equals(
                    NormalizeLaborCode(item.Code),
                    normalized,
                    StringComparison.OrdinalIgnoreCase));
            if (match == null)
            {
                throw new InvalidOperationException(
                    "Chua tao gia nhan cong dieu khien may: " + operatorCode + ".");
            }
            return match;
        }

        private static bool IsCorrosiveMachineContext(
            EstimateV2ResourceRequirement machine)
        {
            if (machine == null)
                return false;
            return machine.NormKeys.Any(key =>
                key.StartsWith("NORM-030.", StringComparison.OrdinalIgnoreCase) ||
                key.StartsWith("NORM-040.", StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeLaborCode(string code)
        {
            string value = (code ?? string.Empty).Trim();
            if (value.StartsWith("LAB-", StringComparison.OrdinalIgnoreCase))
                return value.ToUpperInvariant();
            if (value.StartsWith("bac-", StringComparison.OrdinalIgnoreCase))
            {
                string[] parts = value.Substring(4)
                    .Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 1)
                    return "LAB-QNCN-" + parts[0].ToUpperInvariant();
            }
            return value.ToUpperInvariant();
        }

        private static string LaborCodeFromDescription(string description)
        {
            string value = (description ?? string.Empty).ToLowerInvariant();
            if (value.Contains("5/10"))
                return "LAB-QNCN-5";
            if (value.Contains("7/10"))
                return "LAB-QNCN-7";
            if (value.Contains("8/10"))
                return "LAB-QNCN-8";
            return string.Empty;
        }

        private static bool IsQncnLabor(string code)
        {
            return (code ?? string.Empty).StartsWith(
                "LAB-QNCN-",
                StringComparison.OrdinalIgnoreCase);
        }

        private static string ResourceDisplayName(
            PriceProfile profile,
            string code)
        {
            PriceProfilePrice price;
            if (profile != null && profile.TryFind(code, out price))
                return price.Entry.DisplayName;
            switch ((code ?? string.Empty).ToUpperInvariant())
            {
                case "LAB-QNCN-5": return "Nhân công thợ bậc 5/10";
                case "LAB-QNCN-7": return "Nhân công thợ bậc 7/10";
                case "LAB-QNCN-8": return "Nhân công thợ bậc 8/10";
                default: return code ?? string.Empty;
            }
        }

        private static string FuelDisplayName(
            PriceProfile profile,
            string code)
        {
            PriceProfilePrice price;
            if (profile != null && profile.TryFind(code, out price))
                return price.Entry.DisplayName;
            return code ?? string.Empty;
        }

        private static string ResourceKey(
            NormResourceKind kind,
            string code,
            string unit)
        {
            return ((int)kind).ToString(CultureInfo.InvariantCulture) + "|" +
                (code ?? string.Empty).Trim().ToUpperInvariant() + "|" +
                (unit ?? string.Empty).Trim().ToUpperInvariant();
        }

        private static string HiddenParamReference(
            int metadataStart,
            int row,
            int parameterIndex)
        {
            int column = metadataStart + parameterIndex - 1;
            return ExcelColumnAddress.ToLetters(column) +
                row.ToString(CultureInfo.InvariantCulture);
        }

        private static Excel.Worksheet ResolveOrCreateSheet(
            Excel.Workbook workbook,
            out bool created)
        {
            created = false;
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                Excel.Worksheet nameMatch = null;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index] as Excel.Worksheet;
                        if (sheet == null)
                            continue;

                        WorksheetRole role;
                        if (WorksheetRoleService.TryGetRole(sheet, out role) &&
                            role == WorksheetRole.ResourcePrices)
                        {
                            Excel.Worksheet result = sheet;
                            sheet = null;
                            Release(nameMatch);
                            return result;
                        }

                        if (nameMatch == null &&
                            string.Equals(
                                sheet.Name,
                                SheetName,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            nameMatch = sheet;
                            sheet = null;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }

                if (nameMatch != null)
                    return nameMatch;

                Excel.Worksheet createdSheet = sheets.Add(
                    Type.Missing,
                    sheets.Item[sheets.Count],
                    1,
                    Excel.XlSheetType.xlWorksheet) as Excel.Worksheet;
                if (createdSheet == null)
                    throw new InvalidOperationException("Khong tao duoc sheet VL-NC-M.");
                createdSheet.Name = SheetName;
                created = true;
                return createdSheet;
            }
            finally
            {
                Release(sheets);
            }
        }

        private static void UnmergeVisibleArea(
            Excel.Worksheet sheet,
            int lastRow)
        {
            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    "A1",
                    "F" + Math.Max(1, lastRow).ToString(CultureInfo.InvariantCulture)];
                range.UnMerge();
            }
            finally
            {
                Release(range);
            }
        }

        private static void ClearManagedArea(
            Excel.Worksheet sheet,
            int requiredLastRow,
            int metadataStart,
            ExcelBatchWriteTransaction transaction)
        {
            int lastRow = Math.Max(requiredLastRow, ExistingLastRow(sheet));
            int lastColumn = Math.Max(
                metadataStart + MetadataColumnCount - 1,
                ExistingLastColumn(sheet));
            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    "A1",
                    ExcelColumnAddress.ToLetters(lastColumn) +
                    lastRow.ToString(CultureInfo.InvariantCulture)];
                object[,] empty = new object[lastRow, lastColumn];
                transaction.WriteValue2(range, empty);
            }
            finally
            {
                Release(range);
            }
        }

        private static void WriteMatrix(
            Excel.Worksheet sheet,
            ResourceSheetBuild build,
            ExcelBatchWriteTransaction transaction)
        {
            int columns = build.MetadataStartColumn + MetadataColumnCount - 1;
            object[,] matrix = new object[build.LastRow, columns];

            for (int index = 0; index < build.Rows.Count; index++)
            {
                ResourceSheetRow row = build.Rows[index];
                for (int column = 0; column < VisibleLastColumn; column++)
                    matrix[index, column] = row.Cells[column];

                int meta = build.MetadataStartColumn - 1;
                matrix[index, meta] = row.RowType;
                matrix[index, meta + 1] = row.ResourceCode;
                matrix[index, meta + 2] = row.ResourceKind;
                matrix[index, meta + 3] = row.PackageIdentity;
                matrix[index, meta + 4] = row.InputKey;
                matrix[index, meta + 5] = row.Param1;
                matrix[index, meta + 6] = row.Param2;
                matrix[index, meta + 7] = row.Param3;
                matrix[index, meta + 8] = row.Param4;
                matrix[index, meta + 9] = row.Param5;
                matrix[index, meta + 10] = row.Param6;
            }

            foreach (FuelInputDefinition fuel in build.FuelInputs)
            {
                int row = fuel.HiddenRow - 1;
                int meta = build.MetadataStartColumn - 1;
                matrix[row, meta] = "FUEL_INPUT";
                matrix[row, meta + 1] = fuel.Code;
                matrix[row, meta + 2] = "FuelEnergy";
                matrix[row, meta + 4] = EstimateV2ExcelNames.FuelInput(fuel.Code);
                matrix[row, meta + 5] = fuel.Price.HasValue
                    ? (object)Convert.ToDouble(fuel.Price.Value, CultureInfo.InvariantCulture)
                    : null;
            }

            for (int index = 0; index < MetadataHeaders.Length; index++)
                matrix[0, build.MetadataStartColumn - 1 + index] =
                    MetadataHeaders[index];

            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    "A1",
                    ExcelColumnAddress.ToLetters(columns) +
                    build.LastRow.ToString(CultureInfo.InvariantCulture)];
                transaction.WriteFormula(range, matrix);
            }
            finally
            {
                Release(range);
            }
        }

        private static void ApplyNames(
            Excel.Workbook workbook,
            Excel.Worksheet sheet,
            ResourceSheetBuild build)
        {
            foreach (NamedCell named in build.NamedCells)
                SetWorkbookName(workbook, named.Name, sheet, named.Row, named.Column);
        }

        private static void SetWorkbookName(
            Excel.Workbook workbook,
            string name,
            Excel.Worksheet sheet,
            int row,
            int column)
        {
            Excel.Names names = null;
            Excel.Name existing = null;
            Excel.Name created = null;
            Excel.Range cell = null;
            try
            {
                names = workbook.Names;
                try
                {
                    existing = names.Item(name, Type.Missing, Type.Missing);
                }
                catch (COMException)
                {
                    existing = null;
                }

                cell = sheet.Cells[row, column] as Excel.Range;
                if (cell == null)
                    throw new InvalidOperationException("Khong truy cap duoc o dat Name.");
                string reference = "='" +
                    sheet.Name.Replace("'", "''") +
                    "'!" +
                    cell.Address[
                        true,
                        true,
                        Excel.XlReferenceStyle.xlA1,
                        false,
                        Type.Missing];
                if (existing == null)
                {
                    created = names.Add(
                        name,
                        reference,
                        false,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing);
                }
                else
                {
                    existing.RefersTo = reference;
                }
            }
            finally
            {
                Release(cell);
                Release(created);
                Release(existing);
                Release(names);
            }
        }

        private static decimal? ReadNameDecimal(
            Excel.Workbook workbook,
            string name)
        {
            Excel.Names names = null;
            Excel.Name defined = null;
            Excel.Range range = null;
            try
            {
                try
                {
                    names = workbook.Names;
                    defined = names.Item(name, Type.Missing, Type.Missing);
                }
                catch (COMException)
                {
                    return null;
                }
                range = defined.RefersToRange;
                return DecimalValue(range?.Value2);
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                Release(range);
                Release(defined);
                Release(names);
            }
        }

        private static void FormatSheet(
            Excel.Worksheet sheet,
            ResourceSheetBuild build)
        {
            Excel.Range visible = null;
            Excel.Range title = null;
            Excel.Range all = null;
            try
            {
                visible = sheet.Range[
                    "A1",
                    "F" + build.LastRow.ToString(CultureInfo.InvariantCulture)];
                visible.Font.Name = "Times New Roman";
                visible.Font.Size = 11;
                visible.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                visible.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                visible.Borders.Weight = Excel.XlBorderWeight.xlThin;

                sheet.Columns["A:A"].ColumnWidth = 6;
                sheet.Columns["B:B"].ColumnWidth = 36;
                sheet.Columns["C:C"].ColumnWidth = 12;
                sheet.Columns["D:D"].ColumnWidth = 16;
                sheet.Columns["E:E"].ColumnWidth = 16;
                sheet.Columns["F:F"].ColumnWidth = 18;

                for (int row = 1; row <= build.LastRow; row++)
                {
                    string type = build.Rows[row - 1].RowType;
                    Excel.Range rowRange = null;
                    try
                    {
                        rowRange = sheet.Range[
                            "A" + row.ToString(CultureInfo.InvariantCulture),
                            "F" + row.ToString(CultureInfo.InvariantCulture)];
                        if (type == "TITLE")
                        {
                            rowRange.Merge();
                            rowRange.Font.Bold = true;
                            rowRange.Font.Size = row == 2 ? 14 : 12;
                            rowRange.HorizontalAlignment =
                                Excel.XlHAlign.xlHAlignCenter;
                            rowRange.Interior.Color = ColorRgb(255, 255, 255);
                        }
                        else if (type == "SECTION")
                        {
                            rowRange.Merge();
                            rowRange.Font.Bold = true;
                            rowRange.Interior.Color = ColorRgb(226, 239, 218);
                        }
                        else if (type == "HEADER")
                        {
                            rowRange.Font.Bold = true;
                            rowRange.HorizontalAlignment =
                                Excel.XlHAlign.xlHAlignCenter;
                            rowRange.Interior.Color = ColorRgb(226, 239, 218);
                        }
                        else if (type == "LABOR_HEADER" ||
                                 type == "MACHINE_HEADER")
                        {
                            rowRange.Font.Bold = true;
                            rowRange.Interior.Color = ColorRgb(242, 248, 231);
                        }
                        else if (type == "NOTE")
                        {
                            rowRange.Font.Italic = true;
                        }
                    }
                    finally
                    {
                        Release(rowRange);
                    }
                }

                sheet.Range["C1:F" + build.LastRow.ToString(CultureInfo.InvariantCulture)]
                    .NumberFormat = "#,##0.####";
                sheet.Range["F1:F" + build.LastRow.ToString(CultureInfo.InvariantCulture)]
                    .NumberFormat = "#,##0";
            }
            finally
            {
                Release(all);
                Release(title);
                Release(visible);
            }
        }

        private static int ColorRgb(int red, int green, int blue)
        {
            return red | (green << 8) | (blue << 16);
        }

        private static void HideTechnicalColumns(
            Excel.Worksheet sheet,
            int first,
            int last)
        {
            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    ExcelColumnAddress.ToLetters(first) + ":" +
                    ExcelColumnAddress.ToLetters(last)];
                range.EntireColumn.Hidden = true;
            }
            finally
            {
                Release(range);
            }
        }

        private static void SetWorksheetProperty(
            Excel.Worksheet worksheet,
            string name,
            string value)
        {
            Excel.CustomProperties properties = null;
            Excel.CustomProperty existing = null;
            try
            {
                properties = worksheet.CustomProperties;
                for (int index = properties.Count; index >= 1; index--)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property = properties.Item[index];
                        if (!string.Equals(
                            property.Name,
                            name,
                            StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (existing == null)
                        {
                            existing = property;
                            property = null;
                        }
                        else
                        {
                            property.Delete();
                        }
                    }
                    finally
                    {
                        Release(property);
                    }
                }

                if (existing == null)
                    existing = properties.Add(name, value);
                else
                    existing.Value = value;
            }
            finally
            {
                Release(existing);
                Release(properties);
            }
        }

        private static void DeleteSheetQuietly(
            Excel.Workbook workbook,
            Excel.Worksheet sheet)
        {
            bool oldAlerts = workbook.Application.DisplayAlerts;
            try
            {
                workbook.Application.DisplayAlerts = false;
                sheet.Delete();
            }
            catch
            {
            }
            finally
            {
                workbook.Application.DisplayAlerts = oldAlerts;
            }
        }

        private static int ExistingLastRow(Excel.Worksheet sheet)
        {
            Excel.Range used = null;
            try
            {
                used = sheet.UsedRange;
                return Math.Max(1, used.Row + used.Rows.Count - 1);
            }
            finally
            {
                Release(used);
            }
        }

        private static int ExistingLastColumn(Excel.Worksheet sheet)
        {
            Excel.Range used = null;
            try
            {
                used = sheet.UsedRange;
                return Math.Max(1, used.Column + used.Columns.Count - 1);
            }
            finally
            {
                Release(used);
            }
        }

        private static string ReadText(
            Excel.Worksheet sheet,
            int row,
            int column)
        {
            Excel.Range cell = null;
            try
            {
                cell = sheet.Cells[row, column] as Excel.Range;
                return (Convert.ToString(
                    cell?.Value2,
                    CultureInfo.CurrentCulture) ?? string.Empty).Trim();
            }
            finally
            {
                Release(cell);
            }
        }

        private static decimal? ReadDecimal(
            Excel.Worksheet sheet,
            int row,
            int column)
        {
            Excel.Range cell = null;
            try
            {
                cell = sheet.Cells[row, column] as Excel.Range;
                return DecimalValue(cell?.Value2);
            }
            finally
            {
                Release(cell);
            }
        }

        private static decimal? DecimalValue(object value)
        {
            if (value == null)
                return null;
            try
            {
                return Convert.ToDecimal(value, CultureInfo.CurrentCulture);
            }
            catch
            {
                try
                {
                    return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                }
                catch
                {
                    return null;
                }
            }
        }

        private static string NormalizeText(string value)
        {
            return (value ?? string.Empty)
                .Trim()
                .Replace("  ", " ")
                .ToUpperInvariant();
        }

        private static int Missing(decimal? value)
        {
            return value.HasValue && value.Value != 0m ? 0 : 1;
        }

        private static string Roman(int value)
        {
            if (value <= 0)
                return string.Empty;
            var map = new[]
            {
                new { Value = 1000, Text = "M" },
                new { Value = 900, Text = "CM" },
                new { Value = 500, Text = "D" },
                new { Value = 400, Text = "CD" },
                new { Value = 100, Text = "C" },
                new { Value = 90, Text = "XC" },
                new { Value = 50, Text = "L" },
                new { Value = 40, Text = "XL" },
                new { Value = 10, Text = "X" },
                new { Value = 9, Text = "IX" },
                new { Value = 5, Text = "V" },
                new { Value = 4, Text = "IV" },
                new { Value = 1, Text = "I" }
            };
            int number = value;
            var result = new System.Text.StringBuilder();
            foreach (var item in map)
            {
                while (number >= item.Value)
                {
                    result.Append(item.Text);
                    number -= item.Value;
                }
            }
            return result.ToString();
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private sealed class ResourceSheetBuild
        {
            internal ResourceSheetBuild(
                IList<ResourceSheetRow> rows,
                int lastRow,
                int metadataStartColumn,
                IList<NamedCell> namedCells,
                IEnumerable<FuelInputDefinition> fuelInputs,
                int formulaCount,
                int inputCount,
                int missingInputCount,
                int laborCount)
            {
                Rows = rows;
                LastRow = lastRow;
                MetadataStartColumn = metadataStartColumn;
                NamedCells = namedCells;
                FuelInputs = fuelInputs.ToList();
                FormulaCount = formulaCount;
                InputCount = inputCount;
                MissingInputCount = missingInputCount;
                LaborCount = laborCount;
            }

            internal IList<ResourceSheetRow> Rows { get; }
            internal int LastRow { get; }
            internal int MetadataStartColumn { get; }
            internal IList<NamedCell> NamedCells { get; }
            internal IList<FuelInputDefinition> FuelInputs { get; }
            internal int FormulaCount { get; }
            internal int InputCount { get; }
            internal int MissingInputCount { get; }
            internal int LaborCount { get; }
        }

        private sealed class ResourceSheetRow
        {
            private ResourceSheetRow(
                string rowType,
                object[] cells,
                string resourceCode,
                string resourceKind,
                string packageIdentity,
                string inputKey,
                object param1,
                object param2,
                object param3,
                object param4,
                object param5,
                object param6)
            {
                RowType = rowType ?? string.Empty;
                Cells = cells ?? new object[VisibleLastColumn];
                ResourceCode = resourceCode ?? string.Empty;
                ResourceKind = resourceKind ?? string.Empty;
                PackageIdentity = packageIdentity ?? string.Empty;
                InputKey = inputKey ?? string.Empty;
                Param1 = param1;
                Param2 = param2;
                Param3 = param3;
                Param4 = param4;
                Param5 = param5;
                Param6 = param6;
            }

            internal string RowType { get; }
            internal object[] Cells { get; }
            internal string ResourceCode { get; }
            internal string ResourceKind { get; }
            internal string PackageIdentity { get; }
            internal string InputKey { get; }
            internal object Param1 { get; }
            internal object Param2 { get; }
            internal object Param3 { get; }
            internal object Param4 { get; }
            internal object Param5 { get; }
            internal object Param6 { get; }

            internal static ResourceSheetRow Title(string text)
            {
                return Simple("TITLE", text);
            }

            internal static ResourceSheetRow Section(string text)
            {
                return Simple("SECTION", text);
            }

            internal static ResourceSheetRow Note(string text)
            {
                return Simple("NOTE", text);
            }

            internal static ResourceSheetRow Header(string[] values)
            {
                var cells = new object[VisibleLastColumn];
                for (int index = 0;
                     index < Math.Min(values.Length, VisibleLastColumn);
                     index++)
                    cells[index] = values[index];
                return new ResourceSheetRow(
                    "HEADER", cells, "", "", "", "", null, null, null, null, null, null);
            }

            internal static ResourceSheetRow LaborHeader(
                string roman,
                string title,
                string code,
                string packageIdentity)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = roman;
                cells[1] = title;
                return new ResourceSheetRow(
                    "LABOR_HEADER",
                    cells,
                    code,
                    "Labor",
                    packageIdentity,
                    "",
                    null,
                    null,
                    null,
                    null,
                    null,
                    null);
            }

            internal static ResourceSheetRow LaborInput(
                string stt,
                string title,
                object coefficient,
                object baseSalary,
                object workDays,
                string formula,
                string code,
                string packageIdentity,
                string inputKey)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = stt;
                cells[1] = title;
                cells[2] = coefficient;
                cells[3] = baseSalary;
                cells[4] = workDays;
                cells[5] = formula;
                return new ResourceSheetRow(
                    "LABOR_INPUT",
                    cells,
                    code,
                    "Labor",
                    packageIdentity,
                    inputKey,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null);
            }

            internal static ResourceSheetRow GenericLaborInput(
                string stt,
                string title,
                decimal? inputRate,
                string formula,
                string code,
                string packageIdentity)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = stt;
                cells[1] = title;
                cells[3] = inputRate.HasValue
                    ? (object)Convert.ToDouble(
                        inputRate.Value,
                        CultureInfo.InvariantCulture)
                    : null;
                cells[4] = 1d;
                cells[5] = formula;
                return new ResourceSheetRow(
                    "LABOR_DIRECT_INPUT",
                    cells,
                    code,
                    "Labor",
                    packageIdentity,
                    "DIRECT_DAILY_RATE",
                    null,
                    null,
                    null,
                    null,
                    null,
                    null);
            }

            internal static ResourceSheetRow MachineHeader(
                string roman,
                string title,
                string code,
                string packageIdentity)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = roman;
                cells[1] = title;
                return new ResourceSheetRow(
                    "MACHINE_HEADER",
                    cells,
                    code,
                    "Machine",
                    packageIdentity,
                    "",
                    null,
                    null,
                    null,
                    null,
                    null,
                    null);
            }

            internal static ResourceSheetRow MachineDetail(
                string stt,
                string title,
                string unit,
                decimal quantity,
                string unitPriceFormula,
                string amountFormula,
                string code,
                string packageIdentity,
                string inputKey,
                params object[] parameters)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = stt;
                cells[1] = title;
                cells[2] = unit;
                cells[3] = Convert.ToDouble(quantity, CultureInfo.InvariantCulture);
                cells[4] = unitPriceFormula;
                cells[5] = amountFormula;
                object p1 = parameters.Length > 0 ? parameters[0] : null;
                object p2 = parameters.Length > 1 ? parameters[1] : null;
                object p3 = parameters.Length > 2 ? parameters[2] : null;
                object p4 = parameters.Length > 3 ? parameters[3] : null;
                object p5 = parameters.Length > 4 ? parameters[4] : null;
                object p6 = parameters.Length > 5 ? parameters[5] : null;
                return new ResourceSheetRow(
                    "MACHINE_DETAIL",
                    cells,
                    code,
                    "Machine",
                    packageIdentity,
                    inputKey,
                    p1,
                    p2,
                    p3,
                    p4,
                    p5,
                    p6);
            }

            internal static ResourceSheetRow MaterialInput(
                int index,
                string title,
                string unit,
                decimal? price,
                string code,
                string packageIdentity)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = index;
                cells[1] = title;
                cells[4] = unit;
                cells[5] = price.HasValue
                    ? (object)Convert.ToDouble(
                        price.Value,
                        CultureInfo.InvariantCulture)
                    : null;
                return new ResourceSheetRow(
                    "MATERIAL_INPUT",
                    cells,
                    code,
                    "Material",
                    packageIdentity,
                    "MARKET_PRICE",
                    null,
                    null,
                    null,
                    null,
                    null,
                    null);
            }

            private static ResourceSheetRow Simple(string type, string text)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = text;
                return new ResourceSheetRow(
                    type, cells, "", "", "", "", null, null, null, null, null, null);
            }
        }

        private sealed class PreservedInputs
        {
            internal readonly Dictionary<string, decimal> MaterialPrices =
                new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            internal readonly Dictionary<string, decimal> DirectLaborRates =
                new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            internal readonly Dictionary<string, LaborInputSet> LaborInputs =
                new Dictionary<string, LaborInputSet>(StringComparer.OrdinalIgnoreCase);
            internal readonly Dictionary<string, decimal> FuelPrices =
                new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

            internal LaborInputSet GetLabor(string code, PriceProfile profile)
            {
                LaborInputSet value;
                if (LaborInputs.TryGetValue(code, out value))
                    return value;
                return new LaborInputSet
                {
                    WorkDays = 26m
                };
            }

            internal decimal? GetDirectLaborRate(
                string code,
                PriceProfile profile)
            {
                decimal value;
                if (DirectLaborRates.TryGetValue(code, out value))
                    return value;
                PriceProfilePrice price;
                if (profile != null && profile.TryFind(code, out price))
                    return price.AppliedUnitPriceVnd;
                return null;
            }

            internal decimal? GetMaterialPrice(
                string code,
                PriceProfile profile)
            {
                decimal value;
                if (MaterialPrices.TryGetValue(code, out value))
                    return value;
                PriceProfilePrice price;
                if (profile != null && profile.TryFind(code, out price))
                    return price.AppliedUnitPriceVnd;
                return null;
            }

            internal decimal? GetFuelPrice(
                string code,
                PriceProfile profile)
            {
                decimal value;
                if (FuelPrices.TryGetValue(code, out value))
                    return value;
                PriceProfileEntry entry;
                if (profile != null && profile.TryFind(code, out entry))
                    return profile.FindRequired(code).AppliedUnitPriceVnd;
                return null;
            }
        }

        private sealed class LaborInputSet
        {
            internal decimal? Coefficient;
            internal decimal? BaseSalary;
            internal decimal? WorkDays;
            internal decimal? DangerAllowance;
            internal decimal? MobileAllowance;

            internal bool HasAny =>
                Coefficient.HasValue ||
                BaseSalary.HasValue ||
                WorkDays.HasValue ||
                DangerAllowance.HasValue ||
                MobileAllowance.HasValue;
        }

        private sealed class FuelInputDefinition
        {
            internal FuelInputDefinition(
                string code,
                string unit,
                decimal? price)
            {
                Code = code ?? string.Empty;
                Unit = unit ?? string.Empty;
                Price = price;
            }

            internal string Code { get; }
            internal string Unit { get; }
            internal decimal? Price { get; }
            internal int HiddenRow { get; set; }
        }

        private sealed class NamedCell
        {
            internal NamedCell(
                string name,
                int row,
                int column)
            {
                Name = name;
                Row = row;
                Column = column;
            }

            internal string Name { get; }
            internal int Row { get; }
            internal int Column { get; }
        }
    }

    internal static class MachineFuelRequirementExtensions
    {
        internal static decimal GetAuxiliaryFactorForFormula(
            this MachineFuelRequirement requirement)
        {
            if (requirement == null)
                return 1m;
            switch (requirement.Kind)
            {
                case MachineFuelKind.Diesel: return 1.03m;
                case MachineFuelKind.Gasoline: return 1.02m;
                case MachineFuelKind.Electricity: return 1.05m;
                default: return 1m;
            }
        }
    }
}
