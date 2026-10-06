using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal sealed class EstimateReportsPaneView : EstimateActionPane
    {
        private readonly Excel.Workbook workbook;
        private readonly DataGridView grid;
        internal EstimateReportsPaneView(Excel.Workbook workbook, Action back)
            : base("Báo cáo & Xuất in", EstimateUiIconKind.Document, back)
        {
            this.workbook = workbook;
            Section("Danh sách bảng in");
            grid = new DataGridView { Height = 270, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = System.Drawing.Color.White, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Selected", HeaderText = "In", FillWeight = 20 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Sheet", HeaderText = "Sheet", ReadOnly = true, FillWeight = 65 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Area", HeaderText = "Vùng in", ReadOnly = true, FillWeight = 65 });
            Add(grid);
            Add(Command("Đọc lại danh sách", EstimateUiIconKind.Refresh, () => Run(RefreshSheets), false));
            Add(Command("Mở sheet đã chọn", EstimateUiIconKind.Document, () => Run(OpenSheet), false));
            Add(Command("Kiểm tra", EstimateUiIconKind.Check, () => Run(() => ShowCheck(false)), false));
            Add(Command("Xuất PDF", EstimateUiIconKind.Document, () => Run(Export)));
            Add(Command("Xem trước khi in", EstimateUiIconKind.Document, () => Run(Preview), false));
            Add(Status);
            Run(RefreshSheets);
        }

        private void RefreshSheets()
        {
            var chosen = new HashSet<string>(Selected());
            grid.Rows.Clear();
            foreach (var sheet in WorkbookEstimateV2ReportService.ListSheets(workbook))
            {
                int index = grid.Rows.Add(chosen.Contains(sheet.CodeName), sheet.Name, sheet.PrintArea);
                grid.Rows[index].Tag = sheet.CodeName;
            }
            SetStatus("Đã đọc " + grid.Rows.Count + " sheet.");
        }
        private string[] Selected()
        {
            grid.EndEdit();
            return grid.Rows.Cast<DataGridViewRow>().Where(r => Equals(r.Cells[0].Value, true))
                .Select(r => (string)r.Tag).ToArray();
        }
        private bool ShowCheck(bool confirm)
        {
            var check = WorkbookEstimateV2ReportService.Check(workbook, Selected());
            SetStatus(!check.CanExport ? string.Join(Environment.NewLine, check.Errors) :
                check.Warnings.Count > 0 ? string.Join(Environment.NewLine, check.Warnings) : "Kiểm tra đạt.", !check.CanExport);
            if (!check.CanExport) return false;
            return !confirm || check.Warnings.Count == 0 || MessageBox.Show(this,
                string.Join(Environment.NewLine, check.Warnings) + "\n\nTiếp tục xuất/in?", "Cảnh báo hồ sơ",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes;
        }
        private void Export()
        {
            if (!ShowCheck(true)) return;
            using (var dialog = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", DefaultExt = "pdf", AddExtension = true,
                FileName = System.IO.Path.GetFileNameWithoutExtension(workbook.Name) + ".pdf", OverwritePrompt = true })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                WorkbookEstimateV2ReportService.ExportPdf(workbook, Selected(), dialog.FileName, true);
                SetStatus("Đã xuất PDF: " + dialog.FileName);
            }
        }
        private void Preview()
        {
            if (!ShowCheck(true)) return;
            WorkbookEstimateV2ReportService.PrintPreview(workbook, Selected(), true);
            SetStatus("Đã đóng xem trước.");
        }
        private void OpenSheet()
        {
            if (grid.CurrentRow == null) return;
            Excel.Worksheet sheet = WorkbookEstimateV2ReportService.Resolve(workbook, (string)grid.CurrentRow.Tag);
            try { workbook.Activate(); sheet.Activate(); } finally { Marshal.ReleaseComObject(sheet); }
        }
    }
}
