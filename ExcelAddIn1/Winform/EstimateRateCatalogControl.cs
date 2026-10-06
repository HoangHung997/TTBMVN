using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class EstimateRateCatalogControl : UserControl
    {
        private readonly Excel.Workbook workbook;
        private readonly CultureInfo culture;
        private readonly DataGridView rateGrid;
        private readonly DataGridView resourceGrid;
        private readonly Label statusLabel;
        private WorkbookEstimateWorkspacePreview preview;

        public EstimateRateCatalogControl(Excel.Workbook workbook, CultureInfo culture)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.culture = culture ?? CultureInfo.CurrentCulture;
            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;

            var header = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 48,
                Padding = new Padding(10, 8, 10, 5),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            var refresh = new Button { Text = "Cập nhật", Width = 92, Height = 30 };
            var generate = new Button { Text = "Sinh lại sheet", Width = 108, Height = 30 };
            refresh.Click += (sender, args) => RefreshRates();
            generate.Click += (sender, args) => Generate();
            header.Controls.Add(refresh);
            header.Controls.Add(generate);

            rateGrid = CreateRateGrid();
            resourceGrid = CreateResourceGrid();
            rateGrid.SelectionChanged += (sender, args) => ShowSelectedResources();
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 300,
                Panel1MinSize = 180,
                Panel2MinSize = 150
            };
            split.Panel1.Padding = new Padding(10, 0, 10, 4);
            split.Panel2.Padding = new Padding(10, 4, 10, 4);
            split.Panel1.Controls.Add(rateGrid);
            split.Panel2.Controls.Add(resourceGrid);
            statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 32,
                Padding = new Padding(12, 6, 10, 3),
                AutoEllipsis = true,
                ForeColor = SystemColors.GrayText
            };
            Controls.Add(split);
            Controls.Add(header);
            Controls.Add(statusLabel);
            RefreshRates();
        }

        public void RefreshRates()
        {
            rateGrid.Rows.Clear();
            resourceGrid.Rows.Clear();
            try
            {
                EstimateWorkspace workspace = WorkbookEstimateWorkspaceService.LoadRequired(workbook);
                preview = WorkbookEstimateCalculationService.Preview(workbook, workspace);
                foreach (WorkbookEstimateRatePreview rate in preview.Rates)
                {
                    int index = rateGrid.Rows.Add(
                        rate.IsValid ? "Hợp lệ" : "Lỗi",
                        rate.Group.Identity.RateId,
                        rate.Group.Identity.Environment == EstimateWorkEnvironment.Land ? "Cạn" : "Nước",
                        rate.Group.Identity.LaborAudience == MachineRateAudience.StateBudgetSalary ? "HLNS" : "KHLNS",
                        rate.Group.Identity.NormKey,
                        rate.Group.Identity.VariantCode,
                        rate.Result?.NormTitle ?? string.Empty,
                        rate.Group.Rows.Count,
                        rate.Result?.MaterialAmountVnd,
                        rate.Result?.LaborAmountVnd,
                        rate.Result?.MachineAmountVnd,
                        rate.Result?.TotalAmountVnd);
                    DataGridViewRow row = rateGrid.Rows[index];
                    row.Tag = rate;
                    row.DefaultCellStyle.BackColor = rate.IsValid ? Color.Honeydew : Color.MistyRose;
                    foreach (DataGridViewCell cell in row.Cells)
                        cell.ToolTipText = rate.Error;
                }
                statusLabel.ForeColor = preview.IsValid ? Color.DarkGreen : Color.DarkGoldenrod;
                statusLabel.Text = preview.Rates.Count + " đơn giá duy nhất cho " +
                    preview.CalculatedRowCount + " dòng công tác; " +
                    preview.Plan.TextRows.Count + " dòng văn bản.";
                if (rateGrid.Rows.Count > 0)
                {
                    rateGrid.CurrentCell = rateGrid.Rows[0].Cells["colRateCatalogId"];
                    ShowSelectedResources();
                }
            }
            catch (Exception ex)
            {
                preview = null;
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
            }
        }

        private void Generate()
        {
            try
            {
                RefreshRates();
                if (preview == null || !preview.IsValid)
                    throw new InvalidOperationException("Còn đơn giá lỗi; chưa thể sinh sheet.");
                WorkbookGeneratedEstimateWriteResult result =
                    WorkbookGeneratedEstimateWriter.Apply(workbook, preview);
                workbook.Save();
                statusLabel.ForeColor = Color.DarkGreen;
                statusLabel.Text = "Đã sinh " + result.RateCount + " bảng trên " +
                    result.WorksheetNames.Count + " sheet.";
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Generate rate catalog", "DT-701", "generate", workbook, "UnitRate");
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
                MessageBox.Show(ex.Message, "Không sinh được đơn giá", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ShowSelectedResources()
        {
            resourceGrid.Rows.Clear();
            WorkbookEstimateRatePreview selected = rateGrid.CurrentRow?.Tag as WorkbookEstimateRatePreview;
            if (selected?.Result == null)
                return;
            foreach (UnitRateResourceAmount resource in selected.Result.Resources)
            {
                resourceGrid.Rows.Add(
                    ResourceKind(resource.Kind),
                    resource.ResourceCode,
                    resource.PriceCode,
                    resource.DisplayName,
                    resource.Unit,
                    resource.Quantity,
                    resource.UnitPriceVnd,
                    resource.AmountVnd,
                    resource.IsPriceOverridden,
                    resource.PriceSourceReference);
            }
        }

        private DataGridView CreateRateGrid()
        {
            var grid = Grid();
            grid.Columns.Add(TextColumn("colRateCatalogStatus", "Trạng thái", 75));
            grid.Columns.Add(TextColumn("colRateCatalogId", "Mã đơn giá", 112));
            grid.Columns.Add(TextColumn("colRateCatalogEnvironment", "Cạn/Nước", 70));
            grid.Columns.Add(TextColumn("colRateCatalogAudience", "Đối tượng", 72));
            grid.Columns.Add(TextColumn("colRateCatalogNorm", "Định mức", 105));
            grid.Columns.Add(TextColumn("colRateCatalogVariant", "Mã chi tiết", 94));
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colRateCatalogTitle",
                HeaderText = "Tên công tác",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 180
            });
            grid.Columns.Add(Number("colRateCatalogUses", "Sử dụng", 58, "0"));
            grid.Columns.Add(Number("colRateCatalogMaterial", "VL", 82, "#,##0.##"));
            grid.Columns.Add(Number("colRateCatalogLabor", "NC", 82, "#,##0.##"));
            grid.Columns.Add(Number("colRateCatalogMachine", "Máy", 82, "#,##0.##"));
            grid.Columns.Add(Number("colRateCatalogTotal", "Đơn giá", 100, "#,##0.##"));
            return grid;
        }

        private DataGridView CreateResourceGrid()
        {
            var grid = Grid();
            grid.Columns.Add(TextColumn("colResourceKind", "Loại", 55));
            grid.Columns.Add(TextColumn("colResourceCode", "Mã hao phí", 112));
            grid.Columns.Add(TextColumn("colResourcePriceCode", "Mã giá", 112));
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colResourceName",
                HeaderText = "Tên nguồn lực",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 180
            });
            grid.Columns.Add(TextColumn("colResourceUnit", "ĐVT", 60));
            grid.Columns.Add(Number("colResourceQuantity", "Hao phí", 82, "#,##0.####"));
            grid.Columns.Add(Number("colResourcePrice", "Giá", 92, "#,##0.##"));
            grid.Columns.Add(Number("colResourceAmount", "Thành tiền", 100, "#,##0.##"));
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "colResourceOverride", HeaderText = "Ghi đè", Width = 55 });
            grid.Columns.Add(TextColumn("colResourceSource", "Nguồn", 180));
            return grid;
        }

        private DataGridView Grid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                DefaultCellStyle = { FormatProvider = culture }
            };
        }

        private static DataGridViewTextBoxColumn TextColumn(string name, string title, int width)
        {
            return new DataGridViewTextBoxColumn { Name = name, HeaderText = title, Width = width };
        }

        private static DataGridViewTextBoxColumn Number(string name, string title, int width, string format)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = title,
                Width = width,
                DefaultCellStyle =
                {
                    Alignment = DataGridViewContentAlignment.MiddleRight,
                    Format = format
                }
            };
        }

        private static string ResourceKind(NormResourceKind kind)
        {
            switch (kind)
            {
                case NormResourceKind.Material: return "VL";
                case NormResourceKind.Labor: return "NC";
                case NormResourceKind.Machine: return "M";
                default: return kind.ToString();
            }
        }
    }
}
