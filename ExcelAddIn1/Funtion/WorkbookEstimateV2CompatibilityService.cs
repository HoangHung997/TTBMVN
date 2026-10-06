using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public enum EstimateV2PackageCompatibilityStatus
    {
        NotConfigured = 0,
        Ready = 1,
        MissingOrCorrupt = 2,
        ProfileCorrupt = 3
    }

    public sealed class WorkbookEstimateV2CompatibilityReport
    {
        internal WorkbookEstimateV2CompatibilityReport(
            int assignedRoleCount,
            int migratedSourceCount,
            int renamedSourceCount,
            int legacyAliasCount,
            bool hasDualAudienceLegacyOutputs,
            EstimateV2PackageCompatibilityStatus packageStatus,
            string packageIdentity,
            IEnumerable<string> messages)
        {
            AssignedRoleCount = assignedRoleCount;
            MigratedSourceCount = migratedSourceCount;
            RenamedSourceCount = renamedSourceCount;
            LegacyAliasCount = legacyAliasCount;
            HasDualAudienceLegacyOutputs =
                hasDualAudienceLegacyOutputs;
            PackageStatus = packageStatus;
            PackageIdentity = packageIdentity ?? string.Empty;
            Messages = new ReadOnlyCollection<string>(
                (messages ?? Enumerable.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList());
        }

        public int AssignedRoleCount { get; }
        public int MigratedSourceCount { get; }
        public int RenamedSourceCount { get; }
        public int LegacyAliasCount { get; }
        public bool HasDualAudienceLegacyOutputs { get; }
        public EstimateV2PackageCompatibilityStatus PackageStatus { get; }
        public string PackageIdentity { get; }
        public IReadOnlyList<string> Messages { get; }
        public bool Changed =>
            AssignedRoleCount > 0 ||
            MigratedSourceCount > 0 ||
            RenamedSourceCount > 0;
    }

    public static class WorkbookEstimateV2CompatibilityService
    {
        private static readonly WorksheetRole[] ManagedOutputRoles =
        {
            WorksheetRole.ResourcePrices,
            WorksheetRole.UnitRateLand,
            WorksheetRole.UnitRateWater,
            WorksheetRole.EstimateAppendix,
            WorksheetRole.CostSummary
        };

        public static WorkbookEstimateV2CompatibilityReport Prepare(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var messages = new List<string>();
            List<SheetSnapshot> sheets =
                CaptureSheets(workbook);
            int legacyAliasCount =
                sheets.Count(item =>
                    item.Classification.IsLegacyAlias);
            bool dualAudience =
                HasDualAudienceOutputs(sheets);

            int assignedRoles =
                AssignUnambiguousRoles(
                    workbook,
                    sheets,
                    messages);

            int renamedSources = 0;
            int migratedSources =
                MigrateEstimateSources(
                    workbook,
                    sheets,
                    messages,
                    ref renamedSources);

            EstimateV2PackageCompatibilityStatus packageStatus;
            string packageIdentity;
            InspectPackage(
                workbook,
                out packageStatus,
                out packageIdentity,
                messages);

            return new WorkbookEstimateV2CompatibilityReport(
                assignedRoles,
                migratedSources,
                renamedSources,
                legacyAliasCount,
                dualAudience,
                packageStatus,
                packageIdentity,
                messages);
        }

        public static Excel.Worksheet ResolveOutputWorksheet(
            Excel.Workbook workbook,
            WorksheetRole role,
            params string[] canonicalNames)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            Excel.Worksheet byRole =
                TryResolveUniqueRole(
                    workbook,
                    role);
            if (byRole != null)
                return byRole;

            string[] names = (canonicalNames ??
                new string[0])
                .Where(value =>
                    !string.IsNullOrWhiteSpace(value))
                .ToArray();
            Excel.Worksheet exact =
                FindWorksheetByNames(
                    workbook,
                    names);
            if (exact != null)
                return exact;

            EstimateV2LegacySheetKind kind =
                KindForRole(role);
            if (kind ==
                EstimateV2LegacySheetKind.Unknown)
            {
                return null;
            }

            List<SheetSnapshot> candidates =
                CaptureSheets(workbook)
                    .Where(item =>
                        item.Classification.Kind == kind)
                    .ToList();
            if (candidates.Count != 1)
                return null;

            return ResolveSheet(
                workbook,
                candidates[0].CodeName,
                candidates[0].Name);
        }

        public static Excel.Worksheet ResolveOutputWorksheetByKind(
            Excel.Workbook workbook,
            EstimateV2LegacySheetKind kind,
            params string[] canonicalNames)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            Excel.Worksheet generated =
                FindGeneratedRateSheet(
                    workbook,
                    kind);
            if (generated != null)
                return generated;

            Excel.Worksheet exact =
                FindWorksheetByNames(
                    workbook,
                    canonicalNames);
            if (exact != null)
                return exact;

            List<SheetSnapshot> candidates =
                CaptureSheets(workbook)
                    .Where(item =>
                        item.Classification.Kind == kind)
                    .ToList();
            if (candidates.Count != 1)
                return null;

            return ResolveSheet(
                workbook,
                candidates[0].CodeName,
                candidates[0].Name);
        }

        public static bool NormalizeOutputSheetName(
            Excel.Workbook workbook,
            Excel.Worksheet sheet,
            string canonicalName)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            string wanted =
                (canonicalName ?? string.Empty).Trim();
            if (wanted.Length == 0 ||
                string.Equals(
                    sheet.Name,
                    wanted,
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            Excel.Worksheet existing =
                FindWorksheetByNames(
                    workbook,
                    wanted);
            if (existing != null)
            {
                Release(existing);
                return false;
            }

            try
            {
                sheet.Name = wanted;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        public static int HideSupersededLegacyOutputs(
            Excel.Workbook workbook,
            EstimateV2LegacySheetKind kind,
            string keepCodeName)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            int hidden = 0;
            foreach (SheetSnapshot snapshot in
                CaptureSheets(workbook)
                    .Where(item =>
                        item.Classification.Kind == kind &&
                        item.Classification.IsLegacyAlias))
            {
                if (string.Equals(
                    snapshot.CodeName,
                    keepCodeName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Excel.Worksheet sheet = null;
                try
                {
                    sheet = ResolveSheet(
                        workbook,
                        snapshot.CodeName,
                        snapshot.Name);
                    if (sheet == null)
                        continue;
                    if (sheet.Visible ==
                        Excel.XlSheetVisibility.xlSheetVisible)
                    {
                        sheet.Visible =
                            Excel.XlSheetVisibility.xlSheetHidden;
                        hidden++;
                    }
                }
                catch (COMException)
                {
                }
                finally
                {
                    Release(sheet);
                }
            }

            return hidden;
        }

        private static int MigrateEstimateSources(
            Excel.Workbook workbook,
            IReadOnlyList<SheetSnapshot> snapshots,
            ICollection<string> messages,
            ref int renamedSourceCount)
        {
            IReadOnlyList<EstimateV2RegisteredSource> registered =
                WorkbookEstimateV2RegistrationService
                    .ListRegistered(workbook);
            var registeredKeys =
                new HashSet<string>(
                    registered.Select(item =>
                        item.WorksheetCodeName),
                    StringComparer.OrdinalIgnoreCase);

            List<SheetSnapshot> candidates =
                snapshots
                    .Where(item =>
                        item.Classification.Kind ==
                            EstimateV2LegacySheetKind
                                .EstimateAppendix)
                    .ToList();

            List<SheetSnapshot> canonical =
                candidates
                    .Where(item =>
                        item.Classification.IsCanonicalName)
                    .ToList();

            if (canonical.Count == 1)
            {
                // Khi đã có sheet chuẩn, các sheet Gia DT TC_* còn lại thường
                // là bản audience/copy cũ. Không tự đăng ký để tránh cộng trùng.
                candidates = canonical;
            }
            else if (canonical.Count == 0 &&
                candidates.Count > 1)
            {
                messages.Add(
                    "Phát hiện nhiều bảng Gia DT TC legacy (ví dụ VT/DN/bản sao). " +
                    "Không tự chọn hoặc đăng ký để tránh cộng trùng; hãy chọn đúng bảng ở màn hình Công tác.");
                return 0;
            }
            else if (canonical.Count > 1)
            {
                messages.Add(
                    "Phát hiện nhiều sheet cùng mang tên chuẩn Gia DT TC; không tự migration.");
                return 0;
            }

            if (candidates.Count == 1 &&
                !candidates[0].Classification
                    .IsCanonicalName &&
                FindWorksheetByNames(
                    workbook,
                    "Gia DT TC") == null)
            {
                Excel.Worksheet rename = null;
                try
                {
                    rename = ResolveSheet(
                        workbook,
                        candidates[0].CodeName,
                        candidates[0].Name);
                    if (rename != null)
                    {
                        rename.Name = "Gia DT TC";
                        candidates[0] =
                            new SheetSnapshot(
                                rename.CodeName,
                                rename.Name,
                                EstimateV2CompatibilityRules
                                    .ClassifySheetName(
                                        rename.Name));
                        renamedSourceCount++;
                        messages.Add(
                            "Đã chuẩn hóa tên bảng công tác legacy thành Gia DT TC.");
                    }
                }
                catch (COMException ex)
                {
                    messages.Add(
                        "Không đổi được tên bảng công tác legacy: " +
                        ex.Message);
                }
                finally
                {
                    Release(rename);
                }
            }

            int migrated = 0;
            foreach (SheetSnapshot candidate in candidates)
            {
                if (registeredKeys.Contains(
                    candidate.CodeName))
                {
                    continue;
                }

                Excel.Worksheet sheet = null;
                Excel.Range registerRange = null;
                try
                {
                    sheet = ResolveSheet(
                        workbook,
                        candidate.CodeName,
                        candidate.Name);
                    if (sheet == null)
                        continue;

                    HeaderDetection detection;
                    if (!TryDetectEstimateHeader(
                        sheet,
                        out detection))
                    {
                        messages.Add(
                            sheet.Name +
                            ": phát hiện tên sheet dự toán nhưng không nhận diện được header; giữ nguyên để người dùng đăng ký thủ công.");
                        continue;
                    }

                    if (detection.IsLegacy)
                    {
                        if (!CanConvertLegacyLayout(
                            detection))
                        {
                            messages.Add(
                                sheet.Name +
                                ": layout legacy khác mẫu hỗ trợ; không tự dịch cột để tránh làm sai công thức.");
                            continue;
                        }

                        ConvertLegacyEstimateLayout(
                            sheet,
                            detection,
                            messages);

                        if (!TryDetectEstimateHeader(
                            sheet,
                            out detection) ||
                            detection.IsLegacy)
                        {
                            throw new InvalidOperationException(
                                "Sau migration vẫn không nhận diện được layout V2 trên " +
                                sheet.Name + ".");
                        }
                    }

                    int lastRow =
                        ExistingLastRow(sheet);
                    int firstColumn =
                        Math.Max(
                            1,
                            detection.FirstRequiredColumn);
                    int lastColumn =
                        Math.Max(
                            detection.LastRequiredColumn,
                            detection.QuantityColumn);
                    registerRange = sheet.Range[
                        ExcelColumnAddress.ToLetters(
                            firstColumn) +
                        detection.HeaderRow.ToString(
                            CultureInfo.InvariantCulture),
                        ExcelColumnAddress.ToLetters(
                            lastColumn) +
                        lastRow.ToString(
                            CultureInfo.InvariantCulture)];

                    WorkbookEstimateV2RegistrationService
                        .RegisterSelectedRange(
                            workbook,
                            registerRange);
                    migrated++;
                    registeredKeys.Add(
                        sheet.CodeName);
                    messages.Add(
                        "Đã đăng ký bảng công tác " +
                        sheet.Name +
                        " từ workbook legacy; binding định mức vẫn để người dùng gắn thủ công.");
                }
                catch (Exception ex) when (
                    ex is ArgumentException ||
                    ex is InvalidOperationException ||
                    ex is COMException)
                {
                    messages.Add(
                        candidate.Name +
                        ": migration bảng công tác chưa hoàn tất - " +
                        ex.Message);
                }
                finally
                {
                    Release(registerRange);
                    Release(sheet);
                }
            }

            return migrated;
        }

        private static void ConvertLegacyEstimateLayout(
            Excel.Worksheet sheet,
            HeaderDetection detection,
            ICollection<string> messages)
        {
            int headerRow =
                detection.HeaderRow;
            int lastRow =
                ExistingLastRow(sheet);

            int oldLastColumn =
                ExistingLastColumn(sheet);
            object acceptanceQuantity = null;
            object acceptanceAmounts = null;
            object legacyTail = null;
            Excel.Range acceptanceQuantityRange = null;
            Excel.Range acceptanceAmountRange = null;
            Excel.Range legacyTailRange = null;
            try
            {
                acceptanceQuantityRange = sheet.Range[
                    "E" +
                    headerRow.ToString(
                        CultureInfo.InvariantCulture),
                    "E" +
                    lastRow.ToString(
                        CultureInfo.InvariantCulture)];
                acceptanceQuantity =
                    acceptanceQuantityRange.Value2;

                acceptanceAmountRange = sheet.Range[
                    "L" +
                    headerRow.ToString(
                        CultureInfo.InvariantCulture),
                    "N" +
                    lastRow.ToString(
                        CultureInfo.InvariantCulture)];
                acceptanceAmounts =
                    acceptanceAmountRange.Value2;

                if (oldLastColumn >= 12)
                {
                    legacyTailRange = sheet.Range[
                        "L" +
                        headerRow.ToString(
                            CultureInfo.InvariantCulture),
                        ExcelColumnAddress.ToLetters(
                            oldLastColumn) +
                        lastRow.ToString(
                            CultureInfo.InvariantCulture)];
                    legacyTail =
                        legacyTailRange.Value2;
                }
            }
            finally
            {
                Release(legacyTailRange);
                Release(acceptanceAmountRange);
                Release(acceptanceQuantityRange);
            }

            Excel.Range insert = null;
            Excel.Range delete = null;
            try
            {
                insert = sheet.Range["B:C"];
                insert.Insert(
                    Excel.XlInsertShiftDirection
                        .xlShiftToRight,
                    Type.Missing);

                delete = sheet.Range["G:G"];
                delete.Delete(
                    Excel.XlDeleteShiftDirection
                        .xlShiftToLeft);
            }
            finally
            {
                Release(delete);
                Release(insert);
            }

            WriteText(
                sheet,
                headerRow,
                2,
                "Mã công tác");
            WriteText(
                sheet,
                headerRow,
                3,
                "Định mức");

            int subHeaderRow =
                headerRow + 1;
            if (HasRateSubHeader(
                sheet,
                subHeaderRow))
            {
                MergeHeaderVertical(
                    sheet,
                    2,
                    headerRow,
                    subHeaderRow,
                    "Mã công tác");
                MergeHeaderVertical(
                    sheet,
                    3,
                    headerRow,
                    subHeaderRow,
                    "Định mức");
            }

            RestoreLegacyTailAsValues(
                sheet,
                headerRow,
                lastRow,
                oldLastColumn,
                legacyTail);

            RestoreLegacyAcceptanceSnapshots(
                sheet,
                headerRow,
                lastRow,
                acceptanceQuantity,
                acceptanceAmounts);

            FormatMigratedColumns(
                sheet,
                headerRow,
                lastRow);

            messages.Add(
                sheet.Name +
                ": đã chuyển layout legacy sang A:L chuẩn V2; dữ liệu nghiệm thu cũ được giữ ở cột ẩn ngoài vùng in.");
        }

        private static void RestoreLegacyTailAsValues(
            Excel.Worksheet sheet,
            int headerRow,
            int lastRow,
            int oldLastColumn,
            object legacyTail)
        {
            if (legacyTail == null ||
                oldLastColumn < 12)
            {
                return;
            }

            int columnCount =
                oldLastColumn - 12 + 1;
            int newStart = 13;
            int newEnd =
                newStart + columnCount - 1;
            if (newEnd > 16384)
                return;

            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    ExcelColumnAddress.ToLetters(
                        newStart) +
                    headerRow.ToString(
                        CultureInfo.InvariantCulture),
                    ExcelColumnAddress.ToLetters(
                        newEnd) +
                    lastRow.ToString(
                        CultureInfo.InvariantCulture)];
                range.Value2 = legacyTail;
                range.EntireColumn.Hidden = true;
            }
            finally
            {
                Release(range);
            }
        }

        private static void RestoreLegacyAcceptanceSnapshots(
            Excel.Worksheet sheet,
            int headerRow,
            int lastRow,
            object acceptanceQuantity,
            object acceptanceAmounts)
        {
            int existingLast =
                ExistingLastColumn(sheet);
            int snapshotStart =
                Math.Max(13, existingLast + 1);
            if (snapshotStart + 3 > 16384)
                return;

            Excel.Range quantity = null;
            Excel.Range amounts = null;
            Excel.Range hide = null;
            try
            {
                quantity = sheet.Range[
                    ExcelColumnAddress.ToLetters(
                        snapshotStart) +
                    headerRow.ToString(
                        CultureInfo.InvariantCulture),
                    ExcelColumnAddress.ToLetters(
                        snapshotStart) +
                    lastRow.ToString(
                        CultureInfo.InvariantCulture)];
                quantity.Value2 =
                    acceptanceQuantity;

                amounts = sheet.Range[
                    ExcelColumnAddress.ToLetters(
                        snapshotStart + 1) +
                    headerRow.ToString(
                        CultureInfo.InvariantCulture),
                    ExcelColumnAddress.ToLetters(
                        snapshotStart + 3) +
                    lastRow.ToString(
                        CultureInfo.InvariantCulture)];
                amounts.Value2 =
                    acceptanceAmounts;

                WriteText(
                    sheet,
                    headerRow,
                    snapshotStart,
                    "__LEGACY_NGHIEM_THU");
                WriteText(
                    sheet,
                    headerRow,
                    snapshotStart + 1,
                    "__LEGACY_NT_VL");
                WriteText(
                    sheet,
                    headerRow,
                    snapshotStart + 2,
                    "__LEGACY_NT_NC");
                WriteText(
                    sheet,
                    headerRow,
                    snapshotStart + 3,
                    "__LEGACY_NT_M");

                hide = sheet.Range[
                    "M:" +
                    ExcelColumnAddress.ToLetters(
                        snapshotStart + 3)];
                hide.EntireColumn.Hidden = true;
            }
            finally
            {
                Release(hide);
                Release(amounts);
                Release(quantity);
            }
        }

        private static void FormatMigratedColumns(
            Excel.Worksheet sheet,
            int headerRow,
            int lastRow)
        {
            Excel.Range header = null;
            Excel.Range result = null;
            Excel.Range visible = null;
            try
            {
                header = sheet.Range[
                    "A" +
                    headerRow.ToString(
                        CultureInfo.InvariantCulture),
                    "L" +
                    (headerRow + 1).ToString(
                        CultureInfo.InvariantCulture)];
                header.Font.Bold = true;
                header.HorizontalAlignment =
                    Excel.XlHAlign.xlHAlignCenter;
                header.VerticalAlignment =
                    Excel.XlVAlign.xlVAlignCenter;
                header.WrapText = true;

                result = sheet.Range[
                    "G" +
                    (headerRow + 2).ToString(
                        CultureInfo.InvariantCulture),
                    "L" +
                    lastRow.ToString(
                        CultureInfo.InvariantCulture)];
                result.NumberFormat = "#,##0";

                visible = sheet.Range["A:L"];
                visible.EntireColumn.Hidden = false;

                sheet.Columns["A:A"].ColumnWidth = 7;
                sheet.Columns["B:B"].ColumnWidth = 13;
                sheet.Columns["C:C"].ColumnWidth = 16;
                sheet.Columns["D:D"].ColumnWidth = 48;
                sheet.Columns["E:E"].ColumnWidth = 11;
                sheet.Columns["F:F"].ColumnWidth = 13;
                sheet.Columns["G:L"].ColumnWidth = 14;
            }
            finally
            {
                Release(visible);
                Release(result);
                Release(header);
            }
        }

        private static void MergeHeaderVertical(
            Excel.Worksheet sheet,
            int column,
            int firstRow,
            int secondRow,
            string text)
        {
            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    ExcelColumnAddress.ToLetters(column) +
                    firstRow.ToString(
                        CultureInfo.InvariantCulture),
                    ExcelColumnAddress.ToLetters(column) +
                    secondRow.ToString(
                        CultureInfo.InvariantCulture)];
                range.UnMerge();
                range.Merge();
                range.Value2 = text;
                range.HorizontalAlignment =
                    Excel.XlHAlign.xlHAlignCenter;
                range.VerticalAlignment =
                    Excel.XlVAlign.xlVAlignCenter;
                range.WrapText = true;
            }
            finally
            {
                Release(range);
            }
        }

        private static bool HasRateSubHeader(
            Excel.Worksheet sheet,
            int row)
        {
            string g = Normalize(
                ReadText(sheet, row, 7));
            string h = Normalize(
                ReadText(sheet, row, 8));
            string i = Normalize(
                ReadText(sheet, row, 9));
            return g == "VL" &&
                h == "NC" &&
                i == "M";
        }

        private static bool CanConvertLegacyLayout(
            HeaderDetection detection)
        {
            return detection.HeaderRow > 0 &&
                detection.DescriptionColumn == 2 &&
                detection.UnitColumn == 3 &&
                detection.QuantityColumn == 4 &&
                detection.AcceptanceColumn == 5 &&
                detection.UnitPriceStartColumn == 6 &&
                detection.AmountStartColumn == 9;
        }

        private static bool TryDetectEstimateHeader(
            Excel.Worksheet sheet,
            out HeaderDetection detection)
        {
            int lastRow =
                Math.Min(
                    40,
                    ExistingLastRow(sheet));
            int lastColumn =
                Math.Min(
                    32,
                    ExistingLastColumn(sheet));

            for (int row = 1;
                row <= lastRow;
                row++)
            {
                var headers =
                    new List<string>();
                var columns =
                    new Dictionary<string, int>(
                        StringComparer.OrdinalIgnoreCase);

                for (int column = 1;
                    column <= lastColumn;
                    column++)
                {
                    string value =
                        ReadText(
                            sheet,
                            row,
                            column);
                    if (value.Length == 0)
                        continue;
                    string normalized =
                        Normalize(value);
                    headers.Add(value);

                    if (normalized == "MA CONG TAC")
                        columns["workCode"] = column;
                    else if (normalized == "DINH MUC" ||
                             normalized == "MA DINH MUC" ||
                             normalized == "SO HIEU DINH MUC")
                        columns["norm"] = column;
                    else if (normalized == "MO TA CONG VIEC" ||
                             normalized == "TEN CONG TAC" ||
                             normalized == "NOI DUNG CONG VIEC" ||
                             normalized == "NOI DUNG")
                        columns["description"] = column;
                    else if (normalized == "DON VI" ||
                             normalized == "DVT")
                        columns["unit"] = column;
                    else if (normalized == "KHOI LUONG" ||
                             normalized == "KL")
                        columns["quantity"] = column;
                    else if (normalized.StartsWith(
                        "NGHIEM THU",
                        StringComparison.Ordinal))
                        columns["acceptance"] = column;
                    else if (normalized.StartsWith(
                        "DON GIA",
                        StringComparison.Ordinal))
                        columns["unitPrice"] = column;
                    else if (normalized.StartsWith(
                        "THANH TIEN",
                        StringComparison.Ordinal) &&
                        !columns.ContainsKey("amount"))
                        columns["amount"] = column;
                }

                if (EstimateV2CompatibilityRules
                    .LooksLikeV2EstimateHeader(headers))
                {
                    detection = new HeaderDetection(
                        row,
                        false,
                        Value(columns, "workCode"),
                        Value(columns, "norm"),
                        Value(columns, "description"),
                        Value(columns, "unit"),
                        Value(columns, "quantity"),
                        Value(columns, "acceptance"),
                        Value(columns, "unitPrice"),
                        Value(columns, "amount"));
                    return true;
                }

                if (EstimateV2CompatibilityRules
                    .LooksLikeLegacyEstimateHeader(headers))
                {
                    detection = new HeaderDetection(
                        row,
                        true,
                        0,
                        0,
                        Value(columns, "description"),
                        Value(columns, "unit"),
                        Value(columns, "quantity"),
                        Value(columns, "acceptance"),
                        Value(columns, "unitPrice"),
                        Value(columns, "amount"));
                    return true;
                }
            }

            detection = null;
            return false;
        }

        private static int AssignUnambiguousRoles(
            Excel.Workbook workbook,
            IReadOnlyList<SheetSnapshot> sheets,
            ICollection<string> messages)
        {
            var assigned = new HashSet<WorksheetRole>(
                WorksheetRoleService
                    .ReadAssignments(workbook)
                    .Select(item =>
                    {
                        WorksheetRole role;
                        return WorksheetRoleCatalog
                            .TryParse(item.RoleId, out role)
                                ? (WorksheetRole?)role
                                : null;
                    })
                    .Where(item => item.HasValue)
                    .Select(item => item.Value));

            int count = 0;
            foreach (WorksheetRole role in
                ManagedOutputRoles)
            {
                if (assigned.Contains(role))
                    continue;

                EstimateV2LegacySheetKind kind =
                    KindForRole(role);
                List<SheetSnapshot> candidates =
                    sheets.Where(item =>
                        item.Classification.Kind == kind)
                    .ToList();
                if (candidates.Count == 0)
                    continue;

                List<SheetSnapshot> canonical =
                    candidates.Where(item =>
                        item.Classification.IsCanonicalName)
                    .ToList();
                SheetSnapshot selected =
                    canonical.Count == 1
                        ? canonical[0]
                        : canonical.Count == 0 &&
                          candidates.Count == 1
                            ? candidates[0]
                            : null;
                if (selected == null)
                {
                    messages.Add(
                        WorksheetRoleCatalog.ToId(role) +
                        ": có nhiều sheet legacy ứng viên; không tự chọn sai VT/DN hoặc bản sao.");
                    continue;
                }

                Excel.Worksheet sheet = null;
                try
                {
                    sheet = ResolveSheet(
                        workbook,
                        selected.CodeName,
                        selected.Name);
                    if (sheet == null)
                        continue;
                    WorksheetRoleService.SetRole(
                        sheet,
                        role);
                    count++;
                    assigned.Add(role);
                }
                finally
                {
                    Release(sheet);
                }
            }

            return count;
        }

        private static void InspectPackage(
            Excel.Workbook workbook,
            out EstimateV2PackageCompatibilityStatus status,
            out string identity,
            ICollection<string> messages)
        {
            identity = string.Empty;
            ProjectProfile profile;
            try
            {
                if (!WorkbookProjectProfileService.TryLoad(
                    workbook,
                    out profile))
                {
                    status =
                        EstimateV2PackageCompatibilityStatus
                            .NotConfigured;
                    return;
                }
            }
            catch (Exception ex) when (
                ex is InvalidDataException ||
                ex is ArgumentException ||
                ex is FormatException)
            {
                status =
                    EstimateV2PackageCompatibilityStatus
                        .ProfileCorrupt;
                messages.Add(
                    "Project profile cũ/corrupt: " +
                    ex.Message +
                    ". Module vẫn mở; không tự đổi package.");
                return;
            }

            if (profile == null ||
                string.IsNullOrWhiteSpace(
                    profile.RegulationPackageId))
            {
                status =
                    EstimateV2PackageCompatibilityStatus
                        .NotConfigured;
                return;
            }

            identity =
                profile.RegulationPackageId +
                "@" +
                profile.RegulationPackageVersion +
                "#" +
                profile.RegulationPackageChecksum;

            try
            {
                var store =
                    new RegulationPackageStore(
                        AppPaths.RegulationPackageDirectory);
                store.LoadBundleRequired(
                    profile.RegulationPackageId,
                    profile.RegulationPackageVersion,
                    profile.RegulationPackageChecksum);
                status =
                    EstimateV2PackageCompatibilityStatus
                        .Ready;
            }
            catch (Exception ex) when (
                ex is FileNotFoundException ||
                ex is DirectoryNotFoundException ||
                ex is InvalidDataException ||
                ex is InvalidOperationException ||
                ex is ArgumentException)
            {
                status =
                    EstimateV2PackageCompatibilityStatus
                        .MissingOrCorrupt;
                messages.Add(
                    "Workbook đang pin package " +
                    identity +
                    " nhưng package thiếu/corrupt. Giữ nguyên pin; không tự nâng sang latest. " +
                    ex.Message);
            }
        }

        private static bool HasDualAudienceOutputs(
            IEnumerable<SheetSnapshot> sheets)
        {
            bool state = sheets.Any(item =>
                item.Classification.AudienceHint ==
                    EstimateV2LegacyAudienceHint
                        .StateBudgetSalary);
            bool nonState = sheets.Any(item =>
                item.Classification.AudienceHint ==
                    EstimateV2LegacyAudienceHint
                        .NonStateSalary);
            return state && nonState;
        }

        private static EstimateV2LegacySheetKind KindForRole(
            WorksheetRole role)
        {
            switch (role)
            {
                case WorksheetRole.ResourcePrices:
                    return EstimateV2LegacySheetKind
                        .ResourcePrices;
                case WorksheetRole.UnitRateLand:
                    return EstimateV2LegacySheetKind
                        .UnitRateLand;
                case WorksheetRole.UnitRateWater:
                    return EstimateV2LegacySheetKind
                        .UnitRateWater;
                case WorksheetRole.EstimateAppendix:
                    return EstimateV2LegacySheetKind
                        .EstimateAppendix;
                case WorksheetRole.CostSummary:
                    return EstimateV2LegacySheetKind
                        .CostSummary;
                default:
                    return EstimateV2LegacySheetKind.Unknown;
            }
        }

        private static Excel.Worksheet FindGeneratedRateSheet(
            Excel.Workbook workbook,
            EstimateV2LegacySheetKind kind)
        {
            string expectedEnvironment;
            switch (kind)
            {
                case EstimateV2LegacySheetKind.UnitRateLand:
                    expectedEnvironment =
                        EstimateV2RateEnvironment.Land.ToString();
                    break;
                case EstimateV2LegacySheetKind.UnitRateWater:
                    expectedEnvironment =
                        EstimateV2RateEnvironment.InlandWater.ToString();
                    break;
                case EstimateV2LegacySheetKind.UnitRateSea:
                    expectedEnvironment =
                        EstimateV2RateEnvironment.Sea.ToString();
                    break;
                default:
                    return null;
            }

            Excel.Sheets sheets = null;
            Excel.Worksheet match = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;

                        string environment =
                            ReadWorksheetProperty(
                                sheet,
                                "TTBMVN.EstimateV2.UnitRateEnvironment");
                        if (!string.Equals(
                            environment,
                            expectedEnvironment,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (match != null)
                        {
                            Release(sheet);
                            sheet = null;
                            Release(match);
                            return null;
                        }

                        match = sheet;
                        sheet = null;
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }

                Excel.Worksheet result = match;
                match = null;
                return result;
            }
            finally
            {
                Release(match);
                Release(sheets);
            }
        }

        private static string ReadWorksheetProperty(
            Excel.Worksheet sheet,
            string name)
        {
            Excel.CustomProperties properties = null;
            try
            {
                properties = sheet.CustomProperties;
                for (int index = 1;
                    index <= properties.Count;
                    index++)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property = properties.Item[index];
                        if (string.Equals(
                            property.Name,
                            name,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            return Convert.ToString(
                                property.Value,
                                CultureInfo.InvariantCulture) ??
                                string.Empty;
                        }
                    }
                    finally
                    {
                        Release(property);
                    }
                }
                return string.Empty;
            }
            finally
            {
                Release(properties);
            }
        }

        private static Excel.Worksheet TryResolveUniqueRole(
            Excel.Workbook workbook,
            WorksheetRole role)
        {
            List<WorksheetRoleAssignment> matches =
                WorksheetRoleService
                    .ReadAssignments(workbook)
                    .Where(item =>
                    {
                        WorksheetRole parsed;
                        return WorksheetRoleCatalog
                            .TryParse(
                                item.RoleId,
                                out parsed) &&
                            parsed == role;
                    })
                    .ToList();
            if (matches.Count != 1)
                return null;

            return ResolveSheet(
                workbook,
                matches[0].SheetKey,
                matches[0].SheetName);
        }

        private static Excel.Worksheet FindWorksheetByNames(
            Excel.Workbook workbook,
            params string[] names)
        {
            string[] wanted = (names ??
                new string[0])
                .Where(value =>
                    !string.IsNullOrWhiteSpace(value))
                .ToArray();
            if (wanted.Length == 0)
                return null;

            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;
                        if (wanted.Any(name =>
                            string.Equals(
                                name,
                                sheet.Name,
                                StringComparison.OrdinalIgnoreCase)))
                        {
                            Excel.Worksheet result =
                                sheet;
                            sheet = null;
                            return result;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
                return null;
            }
            finally
            {
                Release(sheets);
            }
        }

        private static Excel.Worksheet ResolveSheet(
            Excel.Workbook workbook,
            string codeName,
            string name)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet fallback = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;
                        if (!string.IsNullOrWhiteSpace(
                                codeName) &&
                            string.Equals(
                                sheet.CodeName,
                                codeName,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Worksheet result =
                                sheet;
                            sheet = null;
                            Release(fallback);
                            return result;
                        }
                        if (fallback == null &&
                            !string.IsNullOrWhiteSpace(name) &&
                            string.Equals(
                                sheet.Name,
                                name,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            fallback = sheet;
                            sheet = null;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
                return fallback;
            }
            finally
            {
                Release(sheets);
            }
        }

        private static List<SheetSnapshot> CaptureSheets(
            Excel.Workbook workbook)
        {
            var result =
                new List<SheetSnapshot>();
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index]
                            as Excel.Worksheet;
                        if (sheet == null)
                            continue;
                        result.Add(
                            new SheetSnapshot(
                                sheet.CodeName,
                                sheet.Name,
                                EstimateV2CompatibilityRules
                                    .ClassifySheetName(
                                        sheet.Name)));
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
                return result;
            }
            finally
            {
                Release(sheets);
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
                    used.Row +
                    used.Rows.Count - 1);
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
                    used.Column +
                    used.Columns.Count - 1);
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
                cell = sheet.Cells[row, column]
                    as Excel.Range;
                return (Convert.ToString(
                    cell?.Value2,
                    CultureInfo.CurrentCulture) ??
                    string.Empty).Trim();
            }
            finally
            {
                Release(cell);
            }
        }

        private static void WriteText(
            Excel.Worksheet sheet,
            int row,
            int column,
            string value)
        {
            Excel.Range cell = null;
            try
            {
                cell = sheet.Cells[row, column]
                    as Excel.Range;
                cell.Value2 = value ?? string.Empty;
            }
            finally
            {
                Release(cell);
            }
        }

        private static int Value(
            IReadOnlyDictionary<string, int> values,
            string key)
        {
            int value;
            return values.TryGetValue(
                key,
                out value)
                ? value
                : 0;
        }

        private static string Normalize(string value)
        {
            return EstimateV2CompatibilityRules
                .Normalize(value);
        }

        private static void Release(
            object value)
        {
            if (value != null &&
                Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }

        private sealed class SheetSnapshot
        {
            internal SheetSnapshot(
                string codeName,
                string name,
                EstimateV2LegacySheetClassification classification)
            {
                CodeName = codeName ?? string.Empty;
                Name = name ?? string.Empty;
                Classification = classification ??
                    throw new ArgumentNullException(
                        nameof(classification));
            }

            internal string CodeName { get; }
            internal string Name { get; }
            internal EstimateV2LegacySheetClassification Classification { get; }
        }

        private sealed class HeaderDetection
        {
            internal HeaderDetection(
                int headerRow,
                bool isLegacy,
                int workCodeColumn,
                int normColumn,
                int descriptionColumn,
                int unitColumn,
                int quantityColumn,
                int acceptanceColumn,
                int unitPriceStartColumn,
                int amountStartColumn)
            {
                HeaderRow = headerRow;
                IsLegacy = isLegacy;
                WorkCodeColumn = workCodeColumn;
                NormColumn = normColumn;
                DescriptionColumn = descriptionColumn;
                UnitColumn = unitColumn;
                QuantityColumn = quantityColumn;
                AcceptanceColumn = acceptanceColumn;
                UnitPriceStartColumn =
                    unitPriceStartColumn;
                AmountStartColumn =
                    amountStartColumn;
            }

            internal int HeaderRow { get; }
            internal bool IsLegacy { get; }
            internal int WorkCodeColumn { get; }
            internal int NormColumn { get; }
            internal int DescriptionColumn { get; }
            internal int UnitColumn { get; }
            internal int QuantityColumn { get; }
            internal int AcceptanceColumn { get; }
            internal int UnitPriceStartColumn { get; }
            internal int AmountStartColumn { get; }

            internal int FirstRequiredColumn =>
                new[]
                {
                    WorkCodeColumn,
                    NormColumn,
                    DescriptionColumn,
                    UnitColumn,
                    QuantityColumn
                }
                .Where(value => value > 0)
                .DefaultIfEmpty(1)
                .Min();

            internal int LastRequiredColumn =>
                new[]
                {
                    WorkCodeColumn,
                    NormColumn,
                    DescriptionColumn,
                    UnitColumn,
                    QuantityColumn
                }.Max();
        }
    }
}
