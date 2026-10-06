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
    public sealed class LegacyPriceApplyResult
    {
        internal LegacyPriceApplyResult(int updated, IEnumerable<string> unmatched)
        {
            Updated = updated;
            UnmatchedLegacyNames = new ReadOnlyCollection<string>(unmatched.ToList());
        }

        public int Updated { get; }
        public IReadOnlyList<string> UnmatchedLegacyNames { get; }
    }

    public static class LegacyWorkbookPriceProfileService
    {
        private sealed class Definition
        {
            internal Definition(
                string legacyName,
                string code,
                PriceResourceKind kind,
                string unit,
                params string[] aliases)
            {
                LegacyName = legacyName;
                Code = code;
                Kind = kind;
                Unit = unit;
                Aliases = aliases ?? Array.Empty<string>();
            }

            internal string LegacyName { get; }
            internal string Code { get; }
            internal PriceResourceKind Kind { get; }
            internal string Unit { get; }
            internal IReadOnlyList<string> Aliases { get; }
        }

        private static readonly IReadOnlyList<Definition> Definitions =
            new ReadOnlyCollection<Definition>(new[]
            {
                new Definition("Nhân công thợ bậc 5/10", "LAB-QNCN-5", PriceResourceKind.Labor, "worker-day", "bac-5-10"),
                new Definition("Nhân công thợ bậc 7/10", "LAB-QNCN-7", PriceResourceKind.Labor, "worker-day", "bac-7-10"),
                new Definition("Nhân công thợ bậc 8/10", "LAB-QNCN-8", PriceResourceKind.Labor, "worker-day", "bac-8-10"),
                new Definition("Máy dò mìn", "MACHINE-M010.001", PriceResourceKind.MachineShift, "shift", "M010.001", "M011.001"),
                new Definition("Máy dò bom trên cạn", "MACHINE-M010.002", PriceResourceKind.MachineShift, "shift", "M010.002", "M011.002"),
                new Definition("Máy dò bom dưới nước", "MACHINE-M010.008", PriceResourceKind.MachineShift, "shift", "M010.008", "M011.008"),
                new Definition("Thuyền cao su tiểu", "MACHINE-M010.009", PriceResourceKind.MachineShift, "shift", "M010.009", "M011.009"),
                new Definition("Thuyền cao su trung", "MACHINE-M010.010", PriceResourceKind.MachineShift, "shift", "M010.010", "M011.010"),
                new Definition("Thuyền Composit", "MACHINE-M010.027", PriceResourceKind.MachineShift, "shift", "M010.027", "M011.027"),
                new Definition("Thiết bị sói và hút bùn cát", "MACHINE-M010.023", PriceResourceKind.MachineShift, "shift", "M010.023", "M011.023"),
                new Definition("Thiết bị lặn sâu 0,5m - 3m", "MACHINE-M010.029", PriceResourceKind.MachineShift, "shift", "M010.029", "M011.029"),
                new Definition("Thiết bị lặn sâu 3m - 6m", "MACHINE-M010.030", PriceResourceKind.MachineShift, "shift", "M010.030", "M011.030"),
                new Definition("Thiết bị lặn sâu 6m - 12m", "MACHINE-M010.031", PriceResourceKind.MachineShift, "shift", "M010.031", "M011.031"),
                new Definition("Cọc bê tông cốt thép (12x12x120)cm", "MAT-CONCRETE-STAKE", PriceResourceKind.Material, "each"),
                new Definition("Cọc gỗ d = 3cm, L= 50 cm", "MAT-WOOD-STAKE", PriceResourceKind.Material, "each"),
                new Definition("Cờ đỏ đuôi nheo", "MAT-RED-FLAG", PriceResourceKind.Material, "each"),
                new Definition("Cờ đỏ to ( 40 x 60 ) cm", "MAT-RED-FLAG-LARGE", PriceResourceKind.Material, "each"),
                new Definition("Dây thừng 10 mm", "MAT-ROPE-10MM", PriceResourceKind.Material, "m"),
                new Definition("Cọc tre d=8, L=200 Cm", "MAT-BAMBOO-STAKE", PriceResourceKind.Material, "each"),
                new Definition("Ván gỗ dày 3 Cm", "MAT-WOOD-BOARD", PriceResourceKind.Material, "m3"),
                new Definition("Đinh 10 cm", "MAT-NAIL-10CM", PriceResourceKind.Material, "kg"),
                new Definition("Đinh 2 mỏ", "LEGACY-MAT-NAIL-TWO-PRONG-EACH", PriceResourceKind.Material, "each"),
                new Definition("Dây thép buộc 2 ly", "MAT-WIRE-2MM", PriceResourceKind.Material, "kg"),
                new Definition("Sào tre d=7, L=500 Cm", "MAT-BAMBOO-POLE", PriceResourceKind.Material, "each"),
                new Definition("Mỏ neo loại 50 kg", "MAT-ANCHOR-50KG", PriceResourceKind.Material, "each"),
                new Definition("Mỏ neo loại 20 kg", "MAT-ANCHOR-20KG", PriceResourceKind.Material, "each"),
                new Definition("Mỏ neo đặc biệt loại 20kg", "MAT-SPECIAL-ANCHOR", PriceResourceKind.Material, "each"),
                new Definition("Phao lớn 1m3", "MAT-LARGE-FLOAT", PriceResourceKind.Material, "each"),
                new Definition("Phao nhựa tròn phi 30 cm", "MAT-PLASTIC-FLOAT", PriceResourceKind.Material, "each"),
                new Definition("Phao nhỏ phi 12 cm", "MAT-SMALL-FLOAT", PriceResourceKind.Material, "each"),
                new Definition("Dây Nilon phi 10mm", "LEGACY-MAT-NYLON-ROPE-10MM", PriceResourceKind.Material, "m"),
                new Definition("Dây Nilon phi 12mm", "MAT-ROPE-12MM", PriceResourceKind.Material, "m"),
                new Definition("Dây Nilon phi 18mm", "MAT-ROPE-18MM", PriceResourceKind.Material, "m"),
                new Definition("Khung gia công bằng tôn 3mm và sắt góc (45x45)mm", "MAT-STEEL-EXCAVATION-FRAME", PriceResourceKind.Material, "kg")
            });

        public static PriceProfile Import(
            Excel.Workbook workbook,
            string location = "Chua khai bao",
            DateTime? createdAtUtc = null)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            ProjectProfile project = WorkbookProjectProfileService.LoadRequired(workbook);
            DateTime valuationDate = (project.PriceDate ?? project.PreparedDate ?? DateTime.Today).Date;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            try
            {
                worksheet = WorksheetRoleService.ResolveWorksheetRequired(workbook, WorksheetRole.ResourcePrices);
                range = worksheet.Range["B6", "F101"];
                Array values = (Array)range.Value2;
                var rows = new Dictionary<string, Tuple<int, decimal>>(StringComparer.OrdinalIgnoreCase);
                for (int offset = 1; offset <= 96; offset++)
                {
                    string name = Convert.ToString(values.GetValue(offset, 1), CultureInfo.CurrentCulture)?.Trim();
                    if (string.IsNullOrEmpty(name))
                        continue;
                    object raw = values.GetValue(offset, 5);
                    decimal price;
                    if (!TryMoney(raw, out price))
                        continue;
                    if (!rows.ContainsKey(name))
                        rows.Add(name, Tuple.Create(offset + 5, price));
                }

                var entries = new List<PriceProfileEntry>();
                var missing = new List<string>();
                foreach (Definition definition in Definitions)
                {
                    Tuple<int, decimal> row;
                    if (!rows.TryGetValue(definition.LegacyName, out row))
                    {
                        missing.Add(definition.LegacyName);
                        continue;
                    }
                    entries.Add(new PriceProfileEntry(
                        definition.Code,
                        definition.Kind,
                        definition.LegacyName,
                        definition.Unit,
                        row.Item2,
                        valuationDate,
                        PriceSourceKind.WorkbookImport,
                        "Workbook:" + worksheet.CodeName + "!F" + row.Item1.ToString(CultureInfo.InvariantCulture),
                        definition.Aliases,
                        definition.LegacyName));
                }
                if (missing.Count > 0)
                    throw new InvalidOperationException("VL-NC-M thieu dong: " + string.Join(", ", missing) + ".");

                return PriceProfile.Create(
                    project.PriceProfileId,
                    "1.0.0",
                    "Nhap tu " + workbook.Name,
                    string.IsNullOrWhiteSpace(location) ? "Chua khai bao" : location.Trim(),
                    valuationDate,
                    MachineRateAudience.NonStateSalary,
                    createdAtUtc ?? DateTime.UtcNow,
                    entries);
            }
            finally
            {
                Release(range);
                Release(worksheet);
            }
        }

        public static LegacyPriceApplyResult ApplyDirectMaterialPrices(
            Excel.Workbook workbook,
            PriceProfile profile)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));
            Excel.Worksheet worksheet = null;
            Excel.Range namesRange = null;
            Excel.Range valuesRange = null;
            try
            {
                worksheet = WorksheetRoleService.ResolveWorksheetRequired(workbook, WorksheetRole.ResourcePrices);
                namesRange = worksheet.Range["B81", "B101"];
                valuesRange = worksheet.Range["F81", "F101"];
                Array names = (Array)namesRange.Value2;
                Array current = (Array)valuesRange.Value2;
                object[,] output = new object[21, 1];
                var byName = profile.Entries
                    .Where(item => item.Kind == PriceResourceKind.Material && item.LegacyLookupName.Length > 0)
                    .ToDictionary(item => item.LegacyLookupName, StringComparer.OrdinalIgnoreCase);
                var matched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int updated = 0;
                for (int index = 1; index <= 21; index++)
                {
                    string name = Convert.ToString(names.GetValue(index, 1), CultureInfo.CurrentCulture)?.Trim();
                    PriceProfileEntry entry;
                    if (!string.IsNullOrEmpty(name) && byName.TryGetValue(name, out entry))
                    {
                        output[index - 1, 0] = profile.FindRequired(entry.Code).AppliedUnitPriceVnd;
                        matched.Add(name);
                        updated++;
                    }
                    else
                    {
                        output[index - 1, 0] = current.GetValue(index, 1);
                    }
                }
                using (new ExcelWriteContext(workbook.Application))
                using (var transaction = new ExcelBatchWriteTransaction())
                {
                    transaction.WriteValue2(valuesRange, output);
                    transaction.Commit();
                }
                workbook.Application.Calculate();
                return new LegacyPriceApplyResult(
                    updated,
                    byName.Keys.Where(name => !matched.Contains(name)).OrderBy(name => name, StringComparer.OrdinalIgnoreCase));
            }
            finally
            {
                Release(valuesRange);
                Release(namesRange);
                Release(worksheet);
            }
        }

        private static bool TryMoney(object raw, out decimal value)
        {
            try
            {
                if (raw == null)
                {
                    value = 0m;
                    return false;
                }
                value = Convert.ToDecimal(raw, CultureInfo.InvariantCulture);
                return value >= 0m;
            }
            catch
            {
                value = 0m;
                return false;
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
