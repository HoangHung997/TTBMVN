using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookEstimateV2RateWriteResult
    {
        internal WorkbookEstimateV2RateWriteResult(
            string worksheetName,
            EstimateV2RateEnvironment environment,
            int rateCount,
            int formulaCount,
            int missingPriceRateCount)
        {
            WorksheetName = worksheetName ?? string.Empty;
            Environment = environment;
            RateCount = rateCount;
            FormulaCount = formulaCount;
            MissingPriceRateCount = missingPriceRateCount;
        }

        public string WorksheetName { get; }
        public EstimateV2RateEnvironment Environment { get; }
        public int RateCount { get; }
        public int FormulaCount { get; }
        public int MissingPriceRateCount { get; }
    }

    public static class WorkbookEstimateV2RateSheetWriter
    {
        private const int VisibleLastColumn = 8; // A:H
        private const int MetadataStartColumn = 9; // I
        private const int MetadataColumnCount = 6;
        private const string GeneratedProperty = "TTBMVN.EstimateV2.UnitRateSheet";
        private const string EnvironmentProperty = "TTBMVN.EstimateV2.UnitRateEnvironment";

        private static readonly string[] MetadataHeaders =
        {
            "__TTB_ROWTYPE",
            "__TTB_RATEID",
            "__TTB_NORM",
            "__TTB_VARIANT",
            "__TTB_PACKAGE",
            "__TTB_ENV"
        };

        public static WorkbookEstimateV2RateWriteResult Apply(
            Excel.Workbook workbook,
            EstimateV2RateEnvironment environment,
            IEnumerable<string> requestedRateIds = null)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            // Tao/cap nhat VL-NC-M truoc de moi cong thuc DG deu co workbook Name
            // on dinh de link, ke ca khi gia hien tai dang trong.
            WorkbookEstimateV2ResourceSheetWriter.Apply(workbook);

            WorkbookEstimateV2RatePreview preview =
                WorkbookEstimateV2RateService.BuildPreview(
                    workbook,
                    environment);
            if (!preview.CanGenerate)
            {
                if (preview.MissingPackageBindings.Count > 0)
                {
                    throw new InvalidOperationException(
                        "Chua resolve du package: " +
                        string.Join(" | ", preview.MissingPackageBindings));
                }
                throw new InvalidOperationException(
                    "Khong co dinh muc " + EnvironmentText(environment) +
                    " nao dang duoc su dung.");
            }

            var requested = new HashSet<string>(
                (requestedRateIds ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim()),
                StringComparer.OrdinalIgnoreCase);

            EstimateV2RateItem[] desired;
            if (requested.Count == 0)
            {
                desired = preview.Items
                    .Select(item => item.Rate)
                    .ToArray();
            }
            else
            {
                desired = preview.Items
                    .Where(item =>
                        requested.Contains(item.Rate.RateId) ||
                        item.IsGenerated)
                    .Select(item => item.Rate)
                    .ToArray();
            }
            if (desired.Length == 0)
                desired = preview.Items.Select(item => item.Rate).ToArray();

            Excel.Worksheet sheet = null;
            bool created = false;
            try
            {
                sheet = ResolveOrCreateSheet(
                    workbook,
                    environment,
                    out created);
                SetWorksheetProperty(
                    sheet,
                    GeneratedProperty,
                    "1");
                SetWorksheetProperty(
                    sheet,
                    EnvironmentProperty,
                    environment.ToString());

                string[] previousRateIds = ReadExistingRateIds(sheet).ToArray();
                IReadOnlyDictionary<string, string> machineNames =
                    BuildMachineNames(
                        workbook,
                        desired);

                RateSheetBuild build = BuildSheet(
                    workbook,
                    desired,
                    environment,
                    machineNames);

                using (new ExcelWriteContext(workbook.Application))
                using (var transaction = new ExcelBatchWriteTransaction())
                {
                    int clearRows = Math.Max(
                        build.LastRow,
                        ExistingLastRow(sheet));
                    int clearColumns = Math.Max(
                        MetadataStartColumn + MetadataColumnCount - 1,
                        ExistingLastColumn(sheet));

                    UnmergeManagedArea(sheet, clearRows);
                    ClearManagedArea(
                        sheet,
                        clearRows,
                        clearColumns,
                        transaction);
                    WriteMatrix(
                        sheet,
                        build,
                        transaction);
                    DeleteRateNames(workbook, previousRateIds);
                    ApplyRateNames(
                        workbook,
                        sheet,
                        build);
                    workbook.Application.Calculate();
                    transaction.Commit();
                }

                FormatSheet(sheet, build);
                HideTechnicalColumns(
                    sheet,
                    MetadataStartColumn,
                    Math.Max(
                        MetadataStartColumn + MetadataColumnCount - 1,
                        ExistingLastColumn(sheet)));
                ConfigurePrint(sheet, build);

                int missingRateCount = preview.Items.Count(item =>
                    desired.Any(rate => string.Equals(
                        rate.RateId,
                        item.Rate.RateId,
                        StringComparison.OrdinalIgnoreCase)) &&
                    item.HasMissingPrice);

                return new WorkbookEstimateV2RateWriteResult(
                    sheet.Name,
                    environment,
                    desired.Length,
                    build.FormulaCount,
                    missingRateCount);
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

        private static RateSheetBuild BuildSheet(
            Excel.Workbook workbook,
            IEnumerable<EstimateV2RateItem> rates,
            EstimateV2RateEnvironment environment,
            IReadOnlyDictionary<string, string> machineNames)
        {
            EstimateV2RateItem[] items = (rates ??
                Enumerable.Empty<EstimateV2RateItem>())
                .OrderBy(item => item.NormCode, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.VariantCode, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var rows = new List<RateSheetRow>();
            var namedCells = new List<NamedCell>();
            int formulaCount = 0;

            rows.Add(RateSheetRow.SheetTitle(
                SheetTitle(environment)));
            rows.Add(RateSheetRow.SheetSubtitle(
                "Các đơn giá được sinh từ định mức đã gắn và liên kết trực tiếp tới VL-NC-M."));
            rows.Add(RateSheetRow.Blank());

            foreach (EstimateV2RateItem rate in items)
            {
                rows.Add(RateSheetRow.RateTitle(
                    rate.Title,
                    rate));
                rows.Add(RateSheetRow.NormCode(
                    "Số hiệu định mức: " +
                    DisplayNorm(rate.NormCode, rate.VariantCode),
                    rate));
                rows.Add(RateSheetRow.WorkUnit(
                    "Đơn vị: " + rate.WorkUnit +
                    "    |    Sử dụng: " +
                    rate.UsageCount.ToString(CultureInfo.InvariantCulture) +
                    " công tác",
                    rate));
                rows.Add(RateSheetRow.Header(rate));

                int materialTotalRow = WriteResourceSection(
                    workbook,
                    rows,
                    rate,
                    NormResourceKind.Material,
                    machineNames,
                    "I",
                    "Vật liệu",
                    6,
                    ref formulaCount);
                int laborTotalRow = WriteResourceSection(
                    workbook,
                    rows,
                    rate,
                    NormResourceKind.Labor,
                    machineNames,
                    "II",
                    "Nhân công",
                    7,
                    ref formulaCount);
                int machineTotalRow = WriteResourceSection(
                    workbook,
                    rows,
                    rate,
                    NormResourceKind.Machine,
                    machineNames,
                    "III",
                    "Máy thi công",
                    8,
                    ref formulaCount);

                int totalRow = rows.Count + 1;
                rows.Add(RateSheetRow.Total(
                    rate,
                    "=SUM(F" + materialTotalRow + ":F" + materialTotalRow + ")",
                    "=SUM(G" + laborTotalRow + ":G" + laborTotalRow + ")",
                    "=SUM(H" + machineTotalRow + ":H" + machineTotalRow + ")"));
                formulaCount += 3;

                int grandTotalRow = rows.Count + 1;
                rows.Add(RateSheetRow.GrandTotal(
                    rate,
                    "=SUM(F" + totalRow + ":H" + totalRow + ")"));
                formulaCount++;

                namedCells.Add(new NamedCell(
                    EstimateV2ExcelNames.RateComponent(rate.RateId, "VL"),
                    totalRow,
                    6));
                namedCells.Add(new NamedCell(
                    EstimateV2ExcelNames.RateComponent(rate.RateId, "NC"),
                    totalRow,
                    7));
                namedCells.Add(new NamedCell(
                    EstimateV2ExcelNames.RateComponent(rate.RateId, "M"),
                    totalRow,
                    8));
                namedCells.Add(new NamedCell(
                    EstimateV2ExcelNames.RateComponent(rate.RateId, "TOTAL"),
                    grandTotalRow,
                    8));

                rows.Add(RateSheetRow.Blank());
            }

            return new RateSheetBuild(
                rows,
                namedCells,
                formulaCount);
        }

        private static int WriteResourceSection(
            Excel.Workbook workbook,
            IList<RateSheetRow> rows,
            EstimateV2RateItem rate,
            NormResourceKind kind,
            IReadOnlyDictionary<string, string> machineNames,
            string roman,
            string title,
            int amountColumn,
            ref int formulaCount)
        {
            rows.Add(RateSheetRow.Section(
                roman,
                title,
                rate));

            EstimateV2RateResource[] resources = rate.Resources
                .Where(item => item.Kind == kind)
                .OrderBy(item => item.IsPercentage)
                .ThenBy(item => item.ResourceCode, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            int firstDirectAmountRow = 0;
            int lastDirectAmountRow = 0;
            int index = 0;
            foreach (EstimateV2RateResource resource in resources
                .Where(item => !item.IsPercentage))
            {
                index++;
                int excelRow = rows.Count + 1;
                if (firstDirectAmountRow == 0)
                    firstDirectAmountRow = excelRow;
                lastDirectAmountRow = excelRow;

                string priceFormula = BuildPriceFormula(
                    resource,
                    rate.PackageIdentity);
                string amountFormula =
                    "=" + ExcelColumnAddress.ToLetters(4) +
                    excelRow.ToString(CultureInfo.InvariantCulture) +
                    "*" +
                    ExcelColumnAddress.ToLetters(5) +
                    excelRow.ToString(CultureInfo.InvariantCulture);

                rows.Add(RateSheetRow.Resource(
                    index,
                    ResolveDisplayName(
                        resource,
                        rate.PackageIdentity,
                        machineNames),
                    resource.Unit,
                    resource.Quantity,
                    priceFormula,
                    amountColumn,
                    amountFormula,
                    rate,
                    resource));
                formulaCount += 2;
            }

            foreach (EstimateV2RateResource resource in resources
                .Where(item => item.IsPercentage))
            {
                index++;
                int excelRow = rows.Count + 1;
                string baseFormula = firstDirectAmountRow > 0
                    ? "SUM(F" +
                        firstDirectAmountRow.ToString(CultureInfo.InvariantCulture) +
                        ":F" +
                        lastDirectAmountRow.ToString(CultureInfo.InvariantCulture) +
                        ")"
                    : "0";
                string amountFormula =
                    "=" + baseFormula + "*D" +
                    excelRow.ToString(CultureInfo.InvariantCulture) + "/100";

                rows.Add(RateSheetRow.Resource(
                    index,
                    "Vật liệu khác",
                    resource.Unit,
                    resource.Quantity,
                    string.Empty,
                    amountColumn,
                    amountFormula,
                    rate,
                    resource));
                formulaCount++;
            }

            int totalRow = rows.Count + 1;
            int firstResourceRow = totalRow - resources.Length;
            string column = ExcelColumnAddress.ToLetters(amountColumn);
            string totalFormula = resources.Length > 0
                ? "=SUM(" + column +
                    firstResourceRow.ToString(CultureInfo.InvariantCulture) +
                    ":" + column +
                    (totalRow - 1).ToString(CultureInfo.InvariantCulture) +
                    ")"
                : "=0";

            rows.Add(RateSheetRow.SectionTotal(
                "Cộng " + ShortKind(kind),
                amountColumn,
                totalFormula,
                rate));
            formulaCount++;
            return totalRow;
        }

        private static string BuildPriceFormula(
            EstimateV2RateResource resource,
            string packageIdentity)
        {
            string[] names = resource.PriceCandidates
                .Select(code => EstimateV2ExcelNames.ResourcePrice(
                    resource.Kind,
                    code,
                    resource.Unit,
                    packageIdentity))
                .ToArray();

            if (names.Length == 0)
                return "=0";
            if (names.Length == 1)
                return "=" + names[0];

            string formula = names[0];
            for (int index = names.Length - 1; index >= 0; index--)
            {
                string fallback = index == names.Length - 1
                    ? names[0]
                    : formula;
                formula =
                    "IF(" + names[index] + ">0," +
                    names[index] + "," + fallback + ")";
            }
            return "=" + formula;
        }

        private static string ResolveDisplayName(
            EstimateV2RateResource resource,
            string packageIdentity,
            IReadOnlyDictionary<string, string> machineNames)
        {
            if (resource.Kind == NormResourceKind.Material ||
                resource.Kind == NormResourceKind.Labor)
            {
                return EstimateV2ResourceNames.Get(resource.ResourceCode);
            }

            if (resource.ResourceCode.EndsWith(
                ".DIVING",
                StringComparison.OrdinalIgnoreCase))
            {
                return "Thiết bị lặn theo độ sâu";
            }

            if (resource.PriceCandidates.Count == 1)
            {
                string key = MachineNameKey(
                    packageIdentity,
                    resource.PriceCandidates[0]);
                string title;
                return machineNames != null &&
                    machineNames.TryGetValue(key, out title)
                    ? title
                    : resource.PriceCandidates[0];
            }

            if (resource.PriceCandidates.Count > 1)
            {
                string[] titles = resource.PriceCandidates
                    .Take(2)
                    .Select(code =>
                    {
                        string key = MachineNameKey(
                            packageIdentity,
                            code);
                        string title;
                        return machineNames != null &&
                            machineNames.TryGetValue(key, out title)
                            ? title
                            : code;
                    })
                    .ToArray();
                return string.Join(" / ", titles);
            }

            return resource.ResourceCode;
        }

        private static IReadOnlyDictionary<string, string> BuildMachineNames(
            Excel.Workbook workbook,
            IEnumerable<EstimateV2RateItem> rates)
        {
            var result = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            EstimateV2State state;
            if (!WorkbookEstimateV2StateService.TryLoad(
                workbook,
                out state))
            {
                return result;
            }

            RegulationPackageBootstrapService.LoadAvailablePackages();
            var store = new RegulationPackageStore(
                AppPaths.RegulationPackageDirectory);

            foreach (string packageIdentity in (rates ??
                Enumerable.Empty<EstimateV2RateItem>())
                .Select(item => item.PackageIdentity)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                EstimateV2WorkItemState sample = state.WorkItems
                    .FirstOrDefault(item => string.Equals(
                        EstimateV2ResourcePlanBuilder.PackageIdentity(item),
                        packageIdentity,
                        StringComparison.OrdinalIgnoreCase));
                if (sample == null)
                    continue;

                try
                {
                    RegulationPackageBundle bundle =
                        store.LoadBundleRequired(
                            sample.PackageId,
                            sample.DataVersion,
                            sample.PackageChecksum);
                    RegulationDataModule module;
                    if (!bundle.Modules.TryGetValue(
                        RegulationModuleKind.MachineRate,
                        out module))
                    {
                        continue;
                    }

                    MachineRateCatalog catalog =
                        MachineRateCatalog.Load(
                            module,
                            MachineRateAudience.NonStateSalary);

                    foreach (string code in (rates ??
                        Enumerable.Empty<EstimateV2RateItem>())
                        .Where(item => string.Equals(
                            item.PackageIdentity,
                            packageIdentity,
                            StringComparison.OrdinalIgnoreCase))
                        .SelectMany(item => item.Machines)
                        .SelectMany(item => item.PriceCandidates)
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct(StringComparer.OrdinalIgnoreCase))
                    {
                        try
                        {
                            MachineRateDefinition definition =
                                ResolveMachine(
                                    catalog,
                                    code);
                            result[MachineNameKey(
                                packageIdentity,
                                code)] =
                                definition.Title;
                        }
                        catch (KeyNotFoundException)
                        {
                        }
                    }
                }
                catch (Exception ex) when (
                    ex is System.IO.IOException ||
                    ex is System.IO.InvalidDataException ||
                    ex is ArgumentException ||
                    ex is KeyNotFoundException)
                {
                    RuntimeLogger.Log(
                        ex,
                        "Resolve V2 DG machine names");
                }
            }

            return result;
        }

        private static MachineRateDefinition ResolveMachine(
            MachineRateCatalog catalog,
            string resourceCode)
        {
            string code = (resourceCode ?? string.Empty).Trim();
            if (code.StartsWith(
                "M010.",
                StringComparison.OrdinalIgnoreCase))
            {
                return catalog.FindRequiredByKey(
                    "MACHINE-" + code.ToUpperInvariant());
            }
            return catalog.FindRequiredByCode(code);
        }

        private static string MachineNameKey(
            string packageIdentity,
            string code)
        {
            return (packageIdentity ?? string.Empty).Trim() +
                "|" +
                (code ?? string.Empty).Trim().ToUpperInvariant();
        }

        private static void WriteMatrix(
            Excel.Worksheet sheet,
            RateSheetBuild build,
            ExcelBatchWriteTransaction transaction)
        {
            int lastColumn =
                MetadataStartColumn + MetadataColumnCount - 1;
            object[,] matrix =
                new object[build.LastRow, lastColumn];

            for (int row = 0; row < build.Rows.Count; row++)
            {
                RateSheetRow item = build.Rows[row];
                for (int column = 0; column < VisibleLastColumn; column++)
                    matrix[row, column] = item.Cells[column];

                matrix[row, MetadataStartColumn - 1] = item.RowType;
                matrix[row, MetadataStartColumn] = item.RateId;
                matrix[row, MetadataStartColumn + 1] = item.NormCode;
                matrix[row, MetadataStartColumn + 2] = item.VariantCode;
                matrix[row, MetadataStartColumn + 3] = item.PackageIdentity;
                matrix[row, MetadataStartColumn + 4] = item.Environment;
            }

            for (int column = 0; column < MetadataHeaders.Length; column++)
                matrix[0, MetadataStartColumn - 1 + column] =
                    MetadataHeaders[column];

            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    "A1",
                    ExcelColumnAddress.ToLetters(lastColumn) +
                    build.LastRow.ToString(CultureInfo.InvariantCulture)];
                transaction.WriteFormula(range, matrix);
            }
            finally
            {
                Release(range);
            }
        }

        private static void ApplyRateNames(
            Excel.Workbook workbook,
            Excel.Worksheet sheet,
            RateSheetBuild build)
        {
            foreach (NamedCell named in build.NamedCells)
                SetWorkbookName(
                    workbook,
                    named.Name,
                    sheet,
                    named.Row,
                    named.Column);
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
                    existing = names.Item(
                        name,
                        Type.Missing,
                        Type.Missing);
                }
                catch (COMException)
                {
                }

                cell = sheet.Cells[row, column] as Excel.Range;
                string reference =
                    "='" + sheet.Name.Replace("'", "''") + "'!" +
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

        private static void DeleteRateNames(
            Excel.Workbook workbook,
            IEnumerable<string> rateIds)
        {
            foreach (string rateId in (rateIds ??
                Enumerable.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                DeleteName(
                    workbook,
                    EstimateV2ExcelNames.RateComponent(rateId, "VL"));
                DeleteName(
                    workbook,
                    EstimateV2ExcelNames.RateComponent(rateId, "NC"));
                DeleteName(
                    workbook,
                    EstimateV2ExcelNames.RateComponent(rateId, "M"));
                DeleteName(
                    workbook,
                    EstimateV2ExcelNames.RateComponent(rateId, "TOTAL"));
            }
        }

        private static void DeleteName(
            Excel.Workbook workbook,
            string name)
        {
            Excel.Names names = null;
            Excel.Name existing = null;
            try
            {
                names = workbook.Names;
                try
                {
                    existing = names.Item(
                        name,
                        Type.Missing,
                        Type.Missing);
                }
                catch (COMException)
                {
                    return;
                }
                existing.Delete();
            }
            finally
            {
                Release(existing);
                Release(names);
            }
        }

        private static IEnumerable<string> ReadExistingRateIds(
            Excel.Worksheet sheet)
        {
            int lastRow = ExistingLastRow(sheet);
            if (lastRow < 2)
                return Enumerable.Empty<string>();

            var result = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            for (int row = 2; row <= lastRow; row++)
            {
                string value = ReadText(
                    sheet,
                    row,
                    MetadataStartColumn + 1);
                if (value.StartsWith(
                    "DG-",
                    StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(value);
                }
            }
            return result.ToArray();
        }

        private static Excel.Worksheet ResolveOrCreateSheet(
            Excel.Workbook workbook,
            EstimateV2RateEnvironment environment,
            out bool created)
        {
            string[] aliases = SheetAliases(environment);
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index] as Excel.Worksheet;
                        if (sheet == null)
                            continue;
                        if (aliases.Any(alias => string.Equals(
                            sheet.Name,
                            alias,
                            StringComparison.OrdinalIgnoreCase)))
                        {
                            Excel.Worksheet result = sheet;
                            sheet = null;
                            created = false;
                            return result;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }

                Excel.Worksheet added =
                    sheets.Add(
                        Type.Missing,
                        sheets.Item[sheets.Count],
                        1,
                        Excel.XlSheetType.xlWorksheet) as Excel.Worksheet;
                added.Name = aliases[0];
                created = true;
                return added;
            }
            finally
            {
                Release(sheets);
            }
        }

        private static string[] SheetAliases(
            EstimateV2RateEnvironment environment)
        {
            switch (environment)
            {
                case EstimateV2RateEnvironment.Land:
                    return new[] { "DG Can", "DG Cạn" };
                case EstimateV2RateEnvironment.InlandWater:
                    return new[] { "DG Nuoc", "DG Nước" };
                case EstimateV2RateEnvironment.Sea:
                    return new[] { "DG Bien", "DG Biển" };
                default:
                    throw new ArgumentOutOfRangeException(nameof(environment));
            }
        }

        private static void ClearManagedArea(
            Excel.Worksheet sheet,
            int lastRow,
            int lastColumn,
            ExcelBatchWriteTransaction transaction)
        {
            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    "A1",
                    ExcelColumnAddress.ToLetters(lastColumn) +
                    lastRow.ToString(CultureInfo.InvariantCulture)];
                transaction.WriteValue2(
                    range,
                    new object[lastRow, lastColumn]);
            }
            finally
            {
                Release(range);
            }
        }

        private static void UnmergeManagedArea(
            Excel.Worksheet sheet,
            int lastRow)
        {
            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    "A1",
                    "H" + lastRow.ToString(CultureInfo.InvariantCulture)];
                range.UnMerge();
            }
            catch (COMException)
            {
            }
            finally
            {
                Release(range);
            }
        }

        private static void FormatSheet(
            Excel.Worksheet sheet,
            RateSheetBuild build)
        {
            Excel.Range visible = null;
            try
            {
                visible = sheet.Range[
                    "A1",
                    "H" + build.LastRow.ToString(CultureInfo.InvariantCulture)];
                visible.Font.Name = "Times New Roman";
                visible.Font.Size = 11;
                visible.VerticalAlignment =
                    Excel.XlVAlign.xlVAlignCenter;
                visible.Borders.LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                visible.Borders.Weight =
                    Excel.XlBorderWeight.xlThin;

                sheet.Columns["A:A"].ColumnWidth = 6;
                sheet.Columns["B:B"].ColumnWidth = 39;
                sheet.Columns["C:C"].ColumnWidth = 12;
                sheet.Columns["D:D"].ColumnWidth = 13;
                sheet.Columns["E:E"].ColumnWidth = 16;
                sheet.Columns["F:H"].ColumnWidth = 16;

                for (int row = 1; row <= build.LastRow; row++)
                {
                    RateSheetRow item = build.Rows[row - 1];
                    Excel.Range rowRange = null;
                    try
                    {
                        rowRange = sheet.Range[
                            "A" + row.ToString(CultureInfo.InvariantCulture),
                            "H" + row.ToString(CultureInfo.InvariantCulture)];

                        switch (item.RowType)
                        {
                            case "SHEET_TITLE":
                                rowRange.Merge();
                                rowRange.Font.Bold = true;
                                rowRange.Font.Size = 14;
                                rowRange.HorizontalAlignment =
                                    Excel.XlHAlign.xlHAlignCenter;
                                rowRange.Interior.Color =
                                    ColorRgb(226, 239, 218);
                                break;
                            case "SHEET_SUBTITLE":
                                rowRange.Merge();
                                rowRange.Font.Italic = true;
                                rowRange.HorizontalAlignment =
                                    Excel.XlHAlign.xlHAlignCenter;
                                break;
                            case "RATE_TITLE":
                                rowRange.Merge();
                                rowRange.Font.Bold = true;
                                rowRange.Interior.Color =
                                    ColorRgb(242, 248, 231);
                                break;
                            case "NORM_CODE":
                            case "WORK_UNIT":
                                rowRange.Merge();
                                break;
                            case "HEADER":
                                rowRange.Font.Bold = true;
                                rowRange.HorizontalAlignment =
                                    Excel.XlHAlign.xlHAlignCenter;
                                rowRange.Interior.Color =
                                    ColorRgb(226, 239, 218);
                                break;
                            case "SECTION":
                                rowRange.Font.Bold = true;
                                rowRange.Interior.Color =
                                    ColorRgb(242, 248, 231);
                                break;
                            case "SECTION_TOTAL":
                            case "TOTAL":
                                rowRange.Font.Bold = true;
                                rowRange.Interior.Color =
                                    ColorRgb(255, 248, 204);
                                break;
                            case "GRAND_TOTAL":
                                rowRange.Font.Bold = true;
                                rowRange.Interior.Color =
                                    ColorRgb(226, 239, 218);
                                break;
                            case "BLANK":
                                rowRange.Borders.LineStyle =
                                    Excel.XlLineStyle.xlLineStyleNone;
                                break;
                        }
                    }
                    finally
                    {
                        Release(rowRange);
                    }
                }

                sheet.Range[
                    "D1:D" + build.LastRow.ToString(CultureInfo.InvariantCulture)]
                    .NumberFormat = "#,##0.######";
                sheet.Range[
                    "E1:H" + build.LastRow.ToString(CultureInfo.InvariantCulture)]
                    .NumberFormat = "#,##0";
                visible.WrapText = true;
            }
            finally
            {
                Release(visible);
            }
        }

        private static void ConfigurePrint(
            Excel.Worksheet sheet,
            RateSheetBuild build)
        {
            sheet.PageSetup.PrintArea =
                "$A$1:$H$" +
                build.LastRow.ToString(CultureInfo.InvariantCulture);
            sheet.PageSetup.Orientation =
                Excel.XlPageOrientation.xlPortrait;
            sheet.PageSetup.BlackAndWhite = true;
            sheet.PageSetup.Zoom = false;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = false;
            sheet.PageSetup.PrintTitleRows = "$1:$1";
        }

        private static void HideTechnicalColumns(
            Excel.Worksheet sheet,
            int first,
            int last)
        {
            if (last < first)
                return;
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

        private static int ExistingLastRow(
            Excel.Worksheet sheet)
        {
            Excel.Range used = null;
            try
            {
                used = sheet.UsedRange;
                return Math.Max(
                    1,
                    used.Row + used.Rows.Count - 1);
            }
            finally
            {
                Release(used);
            }
        }

        private static int ExistingLastColumn(
            Excel.Worksheet sheet)
        {
            Excel.Range used = null;
            try
            {
                used = sheet.UsedRange;
                return Math.Max(
                    1,
                    used.Column + used.Columns.Count - 1);
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
                    CultureInfo.InvariantCulture) ?? string.Empty).Trim();
            }
            finally
            {
                Release(cell);
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

        private static string SheetTitle(
            EstimateV2RateEnvironment environment)
        {
            switch (environment)
            {
                case EstimateV2RateEnvironment.Land:
                    return "PHỤ LỤC CHI TIẾT ĐƠN GIÁ TRÊN CẠN";
                case EstimateV2RateEnvironment.InlandWater:
                    return "PHỤ LỤC CHI TIẾT ĐƠN GIÁ DƯỚI NƯỚC";
                case EstimateV2RateEnvironment.Sea:
                    return "PHỤ LỤC CHI TIẾT ĐƠN GIÁ TRÊN BIỂN";
                default:
                    throw new ArgumentOutOfRangeException(nameof(environment));
            }
        }

        private static string EnvironmentText(
            EstimateV2RateEnvironment environment)
        {
            switch (environment)
            {
                case EstimateV2RateEnvironment.Land:
                    return "trên cạn";
                case EstimateV2RateEnvironment.InlandWater:
                    return "dưới nước";
                case EstimateV2RateEnvironment.Sea:
                    return "trên biển";
                default:
                    return environment.ToString();
            }
        }

        private static string DisplayNorm(
            string normCode,
            string variantCode)
        {
            string code = (normCode ?? string.Empty).Trim();
            if (code.StartsWith(
                "NORM-",
                StringComparison.OrdinalIgnoreCase))
            {
                code = code.Substring(5);
            }
            string variant = (variantCode ?? string.Empty).Trim();
            return variant.Length == 0
                ? code
                : code + " / " + variant;
        }

        private static string ShortKind(NormResourceKind kind)
        {
            switch (kind)
            {
                case NormResourceKind.Material:
                    return "VL";
                case NormResourceKind.Labor:
                    return "NC";
                case NormResourceKind.Machine:
                    return "M";
                default:
                    return kind.ToString();
            }
        }

        private static int ColorRgb(
            int red,
            int green,
            int blue)
        {
            return red | (green << 8) | (blue << 16);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private sealed class RateSheetBuild
        {
            internal RateSheetBuild(
                IList<RateSheetRow> rows,
                IList<NamedCell> namedCells,
                int formulaCount)
            {
                Rows = rows;
                NamedCells = namedCells;
                FormulaCount = formulaCount;
            }

            internal IList<RateSheetRow> Rows { get; }
            internal IList<NamedCell> NamedCells { get; }
            internal int FormulaCount { get; }
            internal int LastRow => Rows.Count;
        }

        private sealed class NamedCell
        {
            internal NamedCell(
                string name,
                int row,
                int column)
            {
                Name = name ?? string.Empty;
                Row = row;
                Column = column;
            }

            internal string Name { get; }
            internal int Row { get; }
            internal int Column { get; }
        }

        private sealed class RateSheetRow
        {
            private RateSheetRow(
                string rowType,
                object[] cells,
                EstimateV2RateItem rate,
                string resourceCode = null)
            {
                RowType = rowType ?? string.Empty;
                Cells = cells ?? new object[VisibleLastColumn];
                RateId = rate?.RateId ?? string.Empty;
                NormCode = rate?.NormCode ?? string.Empty;
                VariantCode = rate?.VariantCode ?? string.Empty;
                PackageIdentity = rate?.PackageIdentity ?? string.Empty;
                Environment = rate == null
                    ? string.Empty
                    : rate.Environment.ToString();
                ResourceCode = resourceCode ?? string.Empty;
            }

            internal string RowType { get; }
            internal object[] Cells { get; }
            internal string RateId { get; }
            internal string NormCode { get; }
            internal string VariantCode { get; }
            internal string PackageIdentity { get; }
            internal string Environment { get; }
            internal string ResourceCode { get; }

            internal static RateSheetRow SheetTitle(string text)
            {
                return Simple("SHEET_TITLE", text);
            }

            internal static RateSheetRow SheetSubtitle(string text)
            {
                return Simple("SHEET_SUBTITLE", text);
            }

            internal static RateSheetRow Blank()
            {
                return Simple("BLANK", string.Empty);
            }

            internal static RateSheetRow RateTitle(
                string title,
                EstimateV2RateItem rate)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = title;
                return new RateSheetRow(
                    "RATE_TITLE",
                    cells,
                    rate);
            }

            internal static RateSheetRow NormCode(
                string text,
                EstimateV2RateItem rate)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = text;
                return new RateSheetRow(
                    "NORM_CODE",
                    cells,
                    rate);
            }

            internal static RateSheetRow WorkUnit(
                string text,
                EstimateV2RateItem rate)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = text;
                return new RateSheetRow(
                    "WORK_UNIT",
                    cells,
                    rate);
            }

            internal static RateSheetRow Header(
                EstimateV2RateItem rate)
            {
                return new RateSheetRow(
                    "HEADER",
                    new object[]
                    {
                        "STT",
                        "Thành phần hao phí",
                        "Đơn vị tính",
                        "Số lượng",
                        "Đơn giá\n(đồng)",
                        "Vật liệu",
                        "Nhân công",
                        "Máy"
                    },
                    rate);
            }

            internal static RateSheetRow Section(
                string roman,
                string title,
                EstimateV2RateItem rate)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = roman;
                cells[1] = title;
                return new RateSheetRow(
                    "SECTION",
                    cells,
                    rate);
            }

            internal static RateSheetRow Resource(
                int index,
                string title,
                string unit,
                decimal quantity,
                string priceFormula,
                int amountColumn,
                string amountFormula,
                EstimateV2RateItem rate,
                EstimateV2RateResource resource)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = index;
                cells[1] = title;
                cells[2] = unit;
                cells[3] = Convert.ToDouble(
                    quantity,
                    CultureInfo.InvariantCulture);
                cells[4] = priceFormula;
                cells[amountColumn - 1] = amountFormula;
                return new RateSheetRow(
                    "RESOURCE",
                    cells,
                    rate,
                    resource.ResourceCode);
            }

            internal static RateSheetRow SectionTotal(
                string title,
                int amountColumn,
                string formula,
                EstimateV2RateItem rate)
            {
                var cells = new object[VisibleLastColumn];
                cells[1] = title;
                cells[amountColumn - 1] = formula;
                return new RateSheetRow(
                    "SECTION_TOTAL",
                    cells,
                    rate);
            }

            internal static RateSheetRow Total(
                EstimateV2RateItem rate,
                string materialFormula,
                string laborFormula,
                string machineFormula)
            {
                var cells = new object[VisibleLastColumn];
                cells[1] = "Cộng:";
                cells[5] = materialFormula;
                cells[6] = laborFormula;
                cells[7] = machineFormula;
                return new RateSheetRow(
                    "TOTAL",
                    cells,
                    rate);
            }

            internal static RateSheetRow GrandTotal(
                EstimateV2RateItem rate,
                string totalFormula)
            {
                var cells = new object[VisibleLastColumn];
                cells[1] = "Tổng cộng đơn giá";
                cells[7] = totalFormula;
                return new RateSheetRow(
                    "GRAND_TOTAL",
                    cells,
                    rate);
            }

            private static RateSheetRow Simple(
                string rowType,
                string text)
            {
                var cells = new object[VisibleLastColumn];
                cells[0] = text;
                return new RateSheetRow(
                    rowType,
                    cells,
                    null);
            }
        }
    }
}
