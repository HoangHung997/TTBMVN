using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class WorkbookEstimateV2CostWriteResult
    {
        internal WorkbookEstimateV2CostWriteResult(
            string thkpWorksheetName,
            int sourceCount,
            int linkedWorkItemCount,
            int missingRateWorkItemCount)
        {
            ThkpWorksheetName = thkpWorksheetName ?? string.Empty;
            SourceCount = sourceCount;
            LinkedWorkItemCount = linkedWorkItemCount;
            MissingRateWorkItemCount = missingRateWorkItemCount;
        }

        public string ThkpWorksheetName { get; }
        public int SourceCount { get; }
        public int LinkedWorkItemCount { get; }
        public int MissingRateWorkItemCount { get; }
    }

    public static class WorkbookEstimateV2CostLinkWriter
    {
        private const string ThkpName = "THKP-TC";
        private const string ThkpHelperProperty =
            "TTBMVN.EstimateV2.ThkpHelperStartColumn";

        public static WorkbookEstimateV2CostWriteResult Apply(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            WorkbookEstimateV2RegistrationService.ReconcileAll(
                workbook);

            EstimateV2State state;
            if (!WorkbookEstimateV2StateService.TryLoad(
                workbook,
                out state))
            {
                throw new InvalidOperationException(
                    "Chua co du lieu cong tac Du toan V2.");
            }

            IReadOnlyList<EstimateV2RegisteredSource> registered =
                WorkbookEstimateV2RegistrationService.ListRegistered(
                    workbook);
            if (registered.Count == 0)
            {
                throw new InvalidOperationException(
                    "Chua dang ky bang cong tac de cap nhat Gia DT TC.");
            }

            var packageErrors = new List<string>();
            EstimateV2CostLinkPlan plan =
                WorkbookEstimateV2CostLinkService.BuildLinkPlan(
                    workbook,
                    state,
                    packageErrors);

            if (packageErrors.Count > 0 &&
                plan.ReadyCount == 0)
            {
                throw new InvalidOperationException(
                    "Khong resolve duoc don gia tu package: " +
                    string.Join(" | ", packageErrors.Take(2)));
            }

            var generatedRates =
                new Dictionary<string, bool>(
                    StringComparer.OrdinalIgnoreCase);
            foreach (string rateId in plan.Links
                .Where(item => item.IsReady)
                .Select(item => item.RateId)
                .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                generatedRates[rateId] =
                    WorkbookEstimateV2RateService.IsGenerated(
                        workbook,
                        rateId);
            }

            Excel.Worksheet thkp = null;
            bool createdThkp = false;
            try
            {
                thkp = ResolveOrCreateThkp(
                    workbook,
                    out createdThkp);

                if (createdThkp)
                    BuildBasicThkpTemplate(thkp);

                ThkpDirectLayout thkpLayout;
                if (!TryDetectThkpLayout(
                    thkp,
                    out thkpLayout))
                {
                    throw new InvalidOperationException(
                        "Khong nhan dien duoc cot Ky hieu/Thanh tien " +
                        "va cac dong VL, NC, M, T tren THKP-TC. " +
                        "Hay dung mau THKP-TC hien hanh hoac sheet V2 moi.");
                }

                var sources =
                    new List<EstimateV2RegisteredSource>();
                foreach (EstimateV2RegisteredSource original in registered)
                {
                    int minimumTechnicalColumn =
                        original.Columns.QuantityColumn + 7;
                    EstimateV2RegisteredSource upgraded =
                        WorkbookEstimateV2RegistrationService
                            .EnsureTechnicalColumnsAfter(
                                workbook,
                                original,
                                minimumTechnicalColumn);
                    sources.Add(upgraded);
                }

                int linked = 0;
                using (new ExcelWriteContext(
                    workbook.Application))
                using (var transaction =
                    new ExcelBatchWriteTransaction())
                {
                    foreach (EstimateV2RegisteredSource source in sources)
                    {
                        Excel.Worksheet sheet = null;
                        Excel.Range outputRange = null;
                        try
                        {
                            sheet =
                                WorkbookEstimateV2CostLinkService
                                    .ResolveWorksheet(
                                        workbook,
                                        source);

                            EnsureCostHeaders(
                                sheet,
                                source);

                            int firstOutput =
                                source.Columns.QuantityColumn + 1;
                            int lastOutput =
                                source.Columns.QuantityColumn + 6;
                            int rowCount =
                                source.LastDataRow -
                                source.FirstDataRow + 1;

                            outputRange = sheet.Range[
                                ExcelColumnAddress.ToLetters(
                                    firstOutput) +
                                source.FirstDataRow.ToString(
                                    CultureInfo.InvariantCulture),
                                ExcelColumnAddress.ToLetters(
                                    lastOutput) +
                                source.LastDataRow.ToString(
                                    CultureInfo.InvariantCulture)];

                            object[,] formulas = CloneMatrix(
                                outputRange.Formula,
                                rowCount,
                                6);

                            for (int row = source.FirstDataRow;
                                row <= source.LastDataRow;
                                row++)
                            {
                                string id =
                                    WorkbookEstimateV2CostLinkService
                                        .ReadText(
                                            sheet,
                                            row,
                                            source.Columns
                                                .TechnicalIdColumn);
                                if (!EstimateV2WorkItemState
                                    .IsValidId(id))
                                {
                                    continue;
                                }

                                int matrixRow =
                                    row - source.FirstDataRow;
                                EstimateV2WorkItemCostLink link =
                                    plan.Find(id);

                                bool canLink =
                                    link != null &&
                                    link.IsReady &&
                                    generatedRates.TryGetValue(
                                        link.RateId,
                                        out bool generated) &&
                                    generated;

                                if (!canLink)
                                {
                                    for (int column = 0;
                                        column < 6;
                                        column++)
                                    {
                                        formulas[matrixRow, column] =
                                            null;
                                    }
                                    continue;
                                }

                                int quantityColumn =
                                    source.Columns.QuantityColumn;
                                int vlPrice = quantityColumn + 1;
                                int ncPrice = quantityColumn + 2;
                                int machinePrice = quantityColumn + 3;
                                int vlAmount = quantityColumn + 4;
                                int ncAmount = quantityColumn + 5;
                                int machineAmount =
                                    quantityColumn + 6;

                                formulas[matrixRow, 0] =
                                    "=" +
                                    EstimateV2ExcelNames
                                        .RateComponent(
                                            link.RateId,
                                            "VL");
                                formulas[matrixRow, 1] =
                                    "=" +
                                    EstimateV2ExcelNames
                                        .RateComponent(
                                            link.RateId,
                                            "NC");
                                formulas[matrixRow, 2] =
                                    "=" +
                                    EstimateV2ExcelNames
                                        .RateComponent(
                                            link.RateId,
                                            "M");

                                string quantity =
                                    ExcelColumnAddress.ToLetters(
                                        quantityColumn) +
                                    row.ToString(
                                        CultureInfo.InvariantCulture);
                                formulas[matrixRow, 3] =
                                    "=" + quantity + "*" +
                                    ExcelColumnAddress.ToLetters(
                                        vlPrice) +
                                    row.ToString(
                                        CultureInfo.InvariantCulture);
                                formulas[matrixRow, 4] =
                                    "=" + quantity + "*" +
                                    ExcelColumnAddress.ToLetters(
                                        ncPrice) +
                                    row.ToString(
                                        CultureInfo.InvariantCulture);
                                formulas[matrixRow, 5] =
                                    "=" + quantity + "*" +
                                    ExcelColumnAddress.ToLetters(
                                        machinePrice) +
                                    row.ToString(
                                        CultureInfo.InvariantCulture);

                                linked++;
                            }

                            transaction.WriteFormula(
                                outputRange,
                                formulas);
                            FormatCostColumns(
                                sheet,
                                source);
                        }
                        finally
                        {
                            WorkbookEstimateV2CostLinkService
                                .Release(outputRange);
                            WorkbookEstimateV2CostLinkService
                                .Release(sheet);
                        }
                    }

                    int helperStart =
                        ResolveThkpHelperStartColumn(thkp);
                    WriteThkpHelpers(
                        thkp,
                        helperStart,
                        sources,
                        transaction);

                    SetWorkbookName(
                        workbook,
                        EstimateV2ExcelNames
                            .EstimateTotal("VL"),
                        thkp,
                        2,
                        helperStart);
                    SetWorkbookName(
                        workbook,
                        EstimateV2ExcelNames
                            .EstimateTotal("NC"),
                        thkp,
                        2,
                        helperStart + 1);
                    SetWorkbookName(
                        workbook,
                        EstimateV2ExcelNames
                            .EstimateTotal("M"),
                        thkp,
                        2,
                        helperStart + 2);
                    SetWorkbookName(
                        workbook,
                        EstimateV2ExcelNames
                            .EstimateTotal("TOTAL"),
                        thkp,
                        2,
                        helperStart + 3);

                    WriteThkpDirectLinks(
                        thkp,
                        thkpLayout,
                        transaction);

                    workbook.Application.Calculate();
                    transaction.Commit();
                }

                ConfigureThkpPrint(thkp, createdThkp);

                int missing =
                    plan.Links.Count(item =>
                        item.Status !=
                            EstimateV2CostLinkStatus.Unbound &&
                        (!item.IsReady ||
                         !generatedRates.TryGetValue(
                             item.RateId,
                             out bool available) ||
                         !available));

                return new WorkbookEstimateV2CostWriteResult(
                    thkp.Name,
                    sources.Count,
                    linked,
                    missing);
            }
            catch
            {
                if (createdThkp && thkp != null)
                    DeleteCreatedSheet(
                        workbook,
                        thkp);
                throw;
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(thkp);
            }
        }

        private static void EnsureCostHeaders(
            Excel.Worksheet sheet,
            EstimateV2RegisteredSource source)
        {
            int q = source.Columns.QuantityColumn;
            int first = q + 1;
            int last = q + 6;
            Excel.Range visibleColumns = null;
            Excel.Range firstGroup = null;
            Excel.Range secondGroup = null;
            Excel.Range headerCell = null;
            Excel.Range secondHeader = null;
            try
            {
                visibleColumns = sheet.Range[
                    ExcelColumnAddress.ToLetters(first) + ":" +
                    ExcelColumnAddress.ToLetters(last)];
                visibleColumns.EntireColumn.Hidden = false;

                bool secondRowIsWorkItem =
                    source.HeaderRow + 1 <=
                        source.LastDataRow &&
                    EstimateV2WorkItemState.IsValidId(
                        WorkbookEstimateV2CostLinkService.ReadText(
                            sheet,
                            source.HeaderRow + 1,
                            source.Columns.TechnicalIdColumn));

                if (secondRowIsWorkItem)
                {
                    string[] oneLevel =
                    {
                        "Đơn giá VL",
                        "Đơn giá NC",
                        "Đơn giá M",
                        "Thành tiền VL",
                        "Thành tiền NC",
                        "Thành tiền M"
                    };
                    for (int index = 0;
                        index < oneLevel.Length;
                        index++)
                    {
                        headerCell = sheet.Cells[
                            source.HeaderRow,
                            first + index] as Excel.Range;
                        headerCell.Value2 =
                            oneLevel[index];
                        FormatHeaderCell(headerCell);
                        WorkbookEstimateV2CostLinkService
                            .Release(headerCell);
                        headerCell = null;
                    }
                }
                else
                {
                    firstGroup = sheet.Range[
                        ExcelColumnAddress.ToLetters(first) +
                        source.HeaderRow.ToString(
                            CultureInfo.InvariantCulture),
                        ExcelColumnAddress.ToLetters(first + 2) +
                        source.HeaderRow.ToString(
                            CultureInfo.InvariantCulture)];
                    firstGroup.UnMerge();
                    firstGroup.Merge();
                    firstGroup.Value2 = "Đơn giá (đồng)";
                    FormatHeaderCell(firstGroup);

                    secondGroup = sheet.Range[
                        ExcelColumnAddress.ToLetters(first + 3) +
                        source.HeaderRow.ToString(
                            CultureInfo.InvariantCulture),
                        ExcelColumnAddress.ToLetters(last) +
                        source.HeaderRow.ToString(
                            CultureInfo.InvariantCulture)];
                    secondGroup.UnMerge();
                    secondGroup.Merge();
                    secondGroup.Value2 =
                        "Thành tiền (đồng)";
                    FormatHeaderCell(secondGroup);

                    string[] subHeaders =
                    {
                        "VL", "NC", "M",
                        "VL", "NC", "M"
                    };
                    for (int index = 0;
                        index < subHeaders.Length;
                        index++)
                    {
                        secondHeader = sheet.Cells[
                            source.HeaderRow + 1,
                            first + index] as Excel.Range;
                        secondHeader.Value2 =
                            subHeaders[index];
                        FormatHeaderCell(secondHeader);
                        WorkbookEstimateV2CostLinkService
                            .Release(secondHeader);
                        secondHeader = null;
                    }
                }
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(secondHeader);
                WorkbookEstimateV2CostLinkService
                    .Release(headerCell);
                WorkbookEstimateV2CostLinkService
                    .Release(secondGroup);
                WorkbookEstimateV2CostLinkService
                    .Release(firstGroup);
                WorkbookEstimateV2CostLinkService
                    .Release(visibleColumns);
            }
        }

        private static void FormatHeaderCell(
            Excel.Range range)
        {
            range.Font.Bold = true;
            range.HorizontalAlignment =
                Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;
            range.WrapText = true;
            range.Interior.Color =
                ColorRgb(226, 239, 218);
            range.Borders.LineStyle =
                Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight =
                Excel.XlBorderWeight.xlThin;
        }

        private static void FormatCostColumns(
            Excel.Worksheet sheet,
            EstimateV2RegisteredSource source)
        {
            int first = source.Columns.QuantityColumn + 1;
            int last = source.Columns.QuantityColumn + 6;
            Excel.Range range = null;
            try
            {
                range = sheet.Range[
                    ExcelColumnAddress.ToLetters(first) +
                    source.FirstDataRow.ToString(
                        CultureInfo.InvariantCulture),
                    ExcelColumnAddress.ToLetters(last) +
                    source.LastDataRow.ToString(
                        CultureInfo.InvariantCulture)];
                range.NumberFormat = "#,##0";
                range.VerticalAlignment =
                    Excel.XlVAlign.xlVAlignCenter;
                range.Borders.LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                range.Borders.Weight =
                    Excel.XlBorderWeight.xlThin;

                for (int column = first;
                    column <= last;
                    column++)
                {
                    Excel.Range whole = null;
                    try
                    {
                        whole = sheet.Range[
                            ExcelColumnAddress.ToLetters(column) +
                            ":" +
                            ExcelColumnAddress.ToLetters(column)];
                        whole.ColumnWidth = 14;
                    }
                    finally
                    {
                        WorkbookEstimateV2CostLinkService
                            .Release(whole);
                    }
                }
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(range);
            }
        }

        private static void WriteThkpHelpers(
            Excel.Worksheet thkp,
            int helperStart,
            IReadOnlyList<EstimateV2RegisteredSource> sources,
            ExcelBatchWriteTransaction transaction)
        {
            object[,] values = new object[2, 4];
            values[0, 0] = "__TTB_GIADT_VL";
            values[0, 1] = "__TTB_GIADT_NC";
            values[0, 2] = "__TTB_GIADT_M";
            values[0, 3] = "__TTB_GIADT_TOTAL";
            values[1, 0] = BuildSourceTotalFormula(
                sources,
                0);
            values[1, 1] = BuildSourceTotalFormula(
                sources,
                1);
            values[1, 2] = BuildSourceTotalFormula(
                sources,
                2);

            string vl =
                ExcelColumnAddress.ToLetters(
                    helperStart) + "2";
            string machine =
                ExcelColumnAddress.ToLetters(
                    helperStart + 2) + "2";
            values[1, 3] =
                "=SUM(" + vl + ":" +
                machine + ")";

            Excel.Range range = null;
            Excel.Range columns = null;
            try
            {
                range = thkp.Range[
                    ExcelColumnAddress.ToLetters(
                        helperStart) + "1",
                    ExcelColumnAddress.ToLetters(
                        helperStart + 3) + "2"];
                transaction.WriteFormula(
                    range,
                    values);

                columns = thkp.Range[
                    ExcelColumnAddress.ToLetters(
                        helperStart) + ":" +
                    ExcelColumnAddress.ToLetters(
                        helperStart + 3)];
                columns.EntireColumn.Hidden = true;
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(columns);
                WorkbookEstimateV2CostLinkService
                    .Release(range);
            }
        }

        private static string BuildSourceTotalFormula(
            IEnumerable<EstimateV2RegisteredSource> sources,
            int componentOffset)
        {
            var parts = new List<string>();
            foreach (EstimateV2RegisteredSource source in sources)
            {
                int amountColumn =
                    source.Columns.QuantityColumn +
                    4 + componentOffset;
                string sheet =
                    "'" +
                    source.WorksheetName.Replace(
                        "'",
                        "''") +
                    "'";
                string first =
                    source.FirstDataRow.ToString(
                        CultureInfo.InvariantCulture);
                string last =
                    source.LastDataRow.ToString(
                        CultureInfo.InvariantCulture);
                string idColumn =
                    ExcelColumnAddress.ToLetters(
                        source.Columns.TechnicalIdColumn);
                string amount =
                    ExcelColumnAddress.ToLetters(
                        amountColumn);

                parts.Add(
                    "SUMIF(" +
                    sheet + "!$" + idColumn + "$" +
                    first + ":$" + idColumn + "$" +
                    last +
                    ",\"<>\"," +
                    sheet + "!$" + amount + "$" +
                    first + ":$" + amount + "$" +
                    last + ")");
            }

            if (parts.Count == 0)
                return "=0";
            if (parts.Count == 1)
                return "=" + parts[0];
            return "=SUM(" +
                string.Join(",", parts) +
                ")";
        }

        private static void WriteThkpDirectLinks(
            Excel.Worksheet sheet,
            ThkpDirectLayout layout,
            ExcelBatchWriteTransaction transaction)
        {
            WriteFormulaCell(
                sheet,
                layout.MaterialRow,
                layout.AmountColumn,
                "=" +
                EstimateV2ExcelNames
                    .EstimateTotal("VL"),
                transaction);
            WriteFormulaCell(
                sheet,
                layout.LaborRow,
                layout.AmountColumn,
                "=" +
                EstimateV2ExcelNames
                    .EstimateTotal("NC"),
                transaction);
            WriteFormulaCell(
                sheet,
                layout.MachineRow,
                layout.AmountColumn,
                "=" +
                EstimateV2ExcelNames
                    .EstimateTotal("M"),
                transaction);
            WriteFormulaCell(
                sheet,
                layout.DirectTotalRow,
                layout.AmountColumn,
                "=" +
                EstimateV2ExcelNames
                    .EstimateTotal("TOTAL"),
                transaction);
        }

        private static void WriteFormulaCell(
            Excel.Worksheet sheet,
            int row,
            int column,
            string formula,
            ExcelBatchWriteTransaction transaction)
        {
            Excel.Range cell = null;
            try
            {
                cell = sheet.Cells[row, column]
                    as Excel.Range;
                transaction.WriteFormula(
                    cell,
                    formula);
                cell.NumberFormat = "#,##0";
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(cell);
            }
        }

        private static Excel.Worksheet ResolveOrCreateThkp(
            Excel.Workbook workbook,
            out bool created)
        {
            Excel.Worksheet existing =
                WorkbookEstimateV2CostLinkService
                    .FindWorksheet(
                        workbook,
                        ThkpName);
            if (existing != null)
            {
                created = false;
                return existing;
            }

            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                Excel.Worksheet createdSheet =
                    sheets.Add(
                        Type.Missing,
                        sheets.Item[sheets.Count],
                        1,
                        Excel.XlSheetType.xlWorksheet)
                    as Excel.Worksheet;
                if (createdSheet == null)
                {
                    throw new InvalidOperationException(
                        "Khong tao duoc THKP-TC.");
                }
                createdSheet.Name = ThkpName;
                created = true;
                return createdSheet;
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(sheets);
            }
        }

        private static void BuildBasicThkpTemplate(
            Excel.Worksheet sheet)
        {
            object[,] values = new object[11, 7];
            values[0, 0] =
                "THẨM ĐỊNH TỔNG HỢP KINH PHÍ - THANH TOÁN CHI PHÍ (THKP-TC)";
            values[1, 0] =
                "Tổng hợp trực tiếp từ các bảng công tác đã đăng ký";
            values[4, 0] = "STT";
            values[4, 1] = "Nội dung chi phí";
            values[4, 2] = "Ký hiệu";
            values[4, 3] = "Cách tính";
            values[4, 4] = "Thành tiền (đồng)";
            values[4, 5] = "Tỷ trọng (%)";
            values[4, 6] = "Ghi chú";

            values[5, 0] = 1;
            values[5, 1] = "CHI PHÍ TRỰC TIẾP";
            values[5, 2] = "T";
            values[5, 3] = "T = VL + NC + M";
            values[5, 6] =
                "Từ các bảng dự toán chi tiết";

            values[6, 0] = "1.1";
            values[6, 1] = "Chi phí vật liệu";
            values[6, 2] = "VL";
            values[6, 3] =
                "Lấy từ bảng dự toán chi tiết";

            values[7, 0] = "1.2";
            values[7, 1] = "Chi phí nhân công";
            values[7, 2] = "NC";
            values[7, 3] =
                "Lấy từ bảng dự toán chi tiết";

            values[8, 0] = "1.3";
            values[8, 1] = "Chi phí máy thi công";
            values[8, 2] = "M";
            values[8, 3] =
                "Lấy từ bảng dự toán chi tiết";

            values[10, 0] = "Ghi chú:";
            values[10, 1] =
                "V2-401 chỉ thiết lập liên kết chi phí trực tiếp. " +
                "Các chi phí khác và thuế tiếp tục theo mẫu pháp lý/thiết lập của hồ sơ.";

            Excel.Range range = null;
            Excel.Range title = null;
            Excel.Range subtitle = null;
            Excel.Range header = null;
            Excel.Range direct = null;
            Excel.Range note = null;
            try
            {
                range = sheet.Range["A1", "G11"];
                range.Value2 = values;
                range.Font.Name = "Times New Roman";
                range.Font.Size = 11;
                range.VerticalAlignment =
                    Excel.XlVAlign.xlVAlignCenter;
                range.WrapText = true;

                title = sheet.Range["A1", "G1"];
                title.Merge();
                title.Font.Bold = true;
                title.Font.Size = 14;
                title.HorizontalAlignment =
                    Excel.XlHAlign.xlHAlignCenter;

                subtitle = sheet.Range["A2", "G2"];
                subtitle.Merge();
                subtitle.Font.Italic = true;
                subtitle.HorizontalAlignment =
                    Excel.XlHAlign.xlHAlignCenter;

                header = sheet.Range["A5", "G5"];
                header.Font.Bold = true;
                header.HorizontalAlignment =
                    Excel.XlHAlign.xlHAlignCenter;
                header.Interior.Color =
                    ColorRgb(226, 239, 218);
                header.Borders.LineStyle =
                    Excel.XlLineStyle.xlContinuous;

                direct = sheet.Range["A6", "G6"];
                direct.Font.Bold = true;
                direct.Interior.Color =
                    ColorRgb(242, 248, 231);

                sheet.Range["A5", "G9"]
                    .Borders.LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sheet.Range["A5", "G9"]
                    .Borders.Weight =
                    Excel.XlBorderWeight.xlThin;

                note = sheet.Range["B11", "G11"];
                note.Merge();
                note.Font.Italic = true;

                sheet.Columns["A:A"].ColumnWidth = 8;
                sheet.Columns["B:B"].ColumnWidth = 31;
                sheet.Columns["C:C"].ColumnWidth = 12;
                sheet.Columns["D:D"].ColumnWidth = 26;
                sheet.Columns["E:E"].ColumnWidth = 18;
                sheet.Columns["F:F"].ColumnWidth = 13;
                sheet.Columns["G:G"].ColumnWidth = 24;
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(note);
                WorkbookEstimateV2CostLinkService
                    .Release(direct);
                WorkbookEstimateV2CostLinkService
                    .Release(header);
                WorkbookEstimateV2CostLinkService
                    .Release(subtitle);
                WorkbookEstimateV2CostLinkService
                    .Release(title);
                WorkbookEstimateV2CostLinkService
                    .Release(range);
            }
        }

        private static bool TryDetectThkpLayout(
            Excel.Worksheet sheet,
            out ThkpDirectLayout layout)
        {
            int lastRow = Math.Min(
                80,
                ExistingLastRow(sheet));
            int lastColumn = Math.Min(
                14,
                ExistingLastColumn(sheet));

            int headerRow = 0;
            int symbolColumn = 0;
            int amountColumn = 0;

            for (int row = 1;
                row <= Math.Min(20, lastRow);
                row++)
            {
                int candidateSymbol = 0;
                int candidateAmount = 0;
                for (int column = 1;
                    column <= lastColumn;
                    column++)
                {
                    string header = Normalize(
                        WorkbookEstimateV2CostLinkService
                            .ReadText(
                                sheet,
                                row,
                                column));
                    if (header == "KY HIEU")
                        candidateSymbol = column;
                    if (header.StartsWith(
                        "THANH TIEN",
                        StringComparison.Ordinal))
                    {
                        candidateAmount = column;
                    }
                }

                if (candidateSymbol > 0 &&
                    candidateAmount > 0)
                {
                    headerRow = row;
                    symbolColumn = candidateSymbol;
                    amountColumn = candidateAmount;
                    break;
                }
            }

            if (headerRow == 0)
            {
                layout = null;
                return false;
            }

            var rows =
                new Dictionary<string, int>(
                    StringComparer.OrdinalIgnoreCase);
            for (int row = headerRow + 1;
                row <= Math.Min(
                    lastRow,
                    headerRow + 40);
                row++)
            {
                string symbol = Normalize(
                    WorkbookEstimateV2CostLinkService
                        .ReadText(
                            sheet,
                            row,
                            symbolColumn));
                if ((symbol == "VL" ||
                     symbol == "NC" ||
                     symbol == "M" ||
                     symbol == "T") &&
                    !rows.ContainsKey(symbol))
                {
                    rows[symbol] = row;
                }
            }

            if (!rows.ContainsKey("VL") ||
                !rows.ContainsKey("NC") ||
                !rows.ContainsKey("M") ||
                !rows.ContainsKey("T"))
            {
                layout = null;
                return false;
            }

            layout = new ThkpDirectLayout(
                headerRow,
                symbolColumn,
                amountColumn,
                rows["VL"],
                rows["NC"],
                rows["M"],
                rows["T"]);
            return true;
        }

        private static int ResolveThkpHelperStartColumn(
            Excel.Worksheet sheet)
        {
            int existing = ReadWorksheetIntProperty(
                sheet,
                ThkpHelperProperty);
            if (existing >= 1 &&
                existing <= 16381)
            {
                return existing;
            }

            int start = Math.Max(
                8,
                ExistingLastColumn(sheet) + 1);
            if (start + 3 > 16384)
            {
                throw new InvalidOperationException(
                    "Khong con du cot cho helper THKP-TC.");
            }

            SetWorksheetProperty(
                sheet,
                ThkpHelperProperty,
                start.ToString(
                    CultureInfo.InvariantCulture));
            return start;
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

                cell = sheet.Cells[row, column]
                    as Excel.Range;
                string reference =
                    "='" +
                    sheet.Name.Replace(
                        "'",
                        "''") +
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
                    existing.RefersTo =
                        reference;
                }
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(cell);
                WorkbookEstimateV2CostLinkService
                    .Release(created);
                WorkbookEstimateV2CostLinkService
                    .Release(existing);
                WorkbookEstimateV2CostLinkService
                    .Release(names);
            }
        }

        private static void ConfigureThkpPrint(
            Excel.Worksheet sheet,
            bool created)
        {
            if (!created)
                return;

            sheet.PageSetup.PrintArea =
                "$A$1:$G$11";
            sheet.PageSetup.Orientation =
                Excel.XlPageOrientation.xlPortrait;
            sheet.PageSetup.BlackAndWhite = true;
            sheet.PageSetup.Zoom = false;
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = false;
        }

        private static object[,] CloneMatrix(
            object source,
            int rows,
            int columns)
        {
            var result =
                new object[rows, columns];
            Array array = source as Array;
            if (array == null)
            {
                if (rows == 1 &&
                    columns == 1)
                {
                    result[0, 0] = source;
                }
                return result;
            }

            int rowBase =
                array.GetLowerBound(0);
            int columnBase =
                array.GetLowerBound(1);
            for (int row = 0;
                row < rows;
                row++)
            {
                for (int column = 0;
                    column < columns;
                    column++)
                {
                    result[row, column] =
                        array.GetValue(
                            rowBase + row,
                            columnBase + column);
                }
            }
            return result;
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
                WorkbookEstimateV2CostLinkService
                    .Release(used);
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
                WorkbookEstimateV2CostLinkService
                    .Release(used);
            }
        }

        private static int ReadWorksheetIntProperty(
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
                        if (!string.Equals(
                            property.Name,
                            name,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        int value;
                        return int.TryParse(
                            Convert.ToString(
                                property.Value,
                                CultureInfo.InvariantCulture),
                            NumberStyles.Integer,
                            CultureInfo.InvariantCulture,
                            out value)
                            ? value
                            : 0;
                    }
                    finally
                    {
                        WorkbookEstimateV2CostLinkService
                            .Release(property);
                    }
                }
                return 0;
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(properties);
            }
        }

        private static void SetWorksheetProperty(
            Excel.Worksheet sheet,
            string name,
            string value)
        {
            Excel.CustomProperties properties = null;
            Excel.CustomProperty first = null;
            try
            {
                properties = sheet.CustomProperties;
                for (int index = properties.Count;
                    index >= 1;
                    index--)
                {
                    Excel.CustomProperty property = null;
                    try
                    {
                        property = properties.Item[index];
                        if (!string.Equals(
                            property.Name,
                            name,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (first == null)
                        {
                            first = property;
                            property = null;
                        }
                        else
                        {
                            property.Delete();
                        }
                    }
                    finally
                    {
                        WorkbookEstimateV2CostLinkService
                            .Release(property);
                    }
                }

                if (first == null)
                    first = properties.Add(
                        name,
                        value);
                else
                    first.Value = value;
            }
            finally
            {
                WorkbookEstimateV2CostLinkService
                    .Release(first);
                WorkbookEstimateV2CostLinkService
                    .Release(properties);
            }
        }

        private static string Normalize(string value)
        {
            string decomposed =
                (value ?? string.Empty)
                    .Normalize(
                        NormalizationForm.FormD);
            var builder =
                new StringBuilder();
            foreach (char ch in decomposed)
            {
                UnicodeCategory category =
                    CharUnicodeInfo.GetUnicodeCategory(ch);
                if (category !=
                    UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(
                        char.ToUpperInvariant(ch));
                }
            }

            string text =
                builder.ToString()
                    .Replace('Đ', 'D')
                    .Replace('đ', 'D');

            return string.Join(
                " ",
                text.Split(
                    new[]
                    {
                        ' ', '\t', '\r', '\n',
                        '-', '_', '/', '(', ')'
                    },
                    StringSplitOptions
                        .RemoveEmptyEntries));
        }

        private static int ColorRgb(
            int red,
            int green,
            int blue)
        {
            return red |
                (green << 8) |
                (blue << 16);
        }

        private static void DeleteCreatedSheet(
            Excel.Workbook workbook,
            Excel.Worksheet sheet)
        {
            bool oldAlerts =
                workbook.Application.DisplayAlerts;
            try
            {
                workbook.Application.DisplayAlerts =
                    false;
                sheet.Delete();
            }
            catch
            {
            }
            finally
            {
                workbook.Application.DisplayAlerts =
                    oldAlerts;
            }
        }

        private sealed class ThkpDirectLayout
        {
            internal ThkpDirectLayout(
                int headerRow,
                int symbolColumn,
                int amountColumn,
                int materialRow,
                int laborRow,
                int machineRow,
                int directTotalRow)
            {
                HeaderRow = headerRow;
                SymbolColumn = symbolColumn;
                AmountColumn = amountColumn;
                MaterialRow = materialRow;
                LaborRow = laborRow;
                MachineRow = machineRow;
                DirectTotalRow = directTotalRow;
            }

            internal int HeaderRow { get; }
            internal int SymbolColumn { get; }
            internal int AmountColumn { get; }
            internal int MaterialRow { get; }
            internal int LaborRow { get; }
            internal int MachineRow { get; }
            internal int DirectTotalRow { get; }
        }
    }
}
