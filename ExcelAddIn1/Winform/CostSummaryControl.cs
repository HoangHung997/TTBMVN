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
    public sealed class CostSummaryControl : UserControl
    {
        private readonly Excel.Workbook workbook;
        private readonly CultureInfo culture;
        private readonly WorkbookCostSummaryContext context;
        private readonly ComboBox templateInput;
        private readonly ComboBox terrainInput;
        private readonly NumericUpDown areaInput;
        private readonly ComboBox projectInput;
        private readonly ComboBox constructionInput;
        private readonly ComboBox disposalInput;
        private readonly NumericUpDown preTaxIncomeInput;
        private readonly NumericUpDown vatInput;
        private readonly DataGridView componentGrid;
        private readonly Label directValue;
        private readonly Label zValue;
        private readonly Label otherValue;
        private readonly Label beforeTaxValue;
        private readonly Label vatValue;
        private readonly Label finalValue;
        private readonly Label statusLabel;

        public CostSummaryControl(Excel.Workbook workbook, CultureInfo culture)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.culture = culture ?? CultureInfo.CurrentCulture;
            context = WorkbookCostSummaryService.Load(workbook);
            Dock = DockStyle.Fill;
            BackColor = SystemColors.Control;
            Padding = new Padding(10);

            templateInput = CreateComboBox("cboCostSummaryTemplate");
            terrainInput = CreateComboBox("cboCostSummaryTerrain");
            areaInput = CreateNumber("numCostSummaryArea", 3, 0.01m, 1000000m);
            projectInput = CreateComboBox("cboCostSummaryProject");
            constructionInput = CreateComboBox("cboCostSummaryConstruction");
            disposalInput = CreateComboBox("cboCostSummaryDisposal");
            preTaxIncomeInput = CreateNumber("numCostSummaryPreTaxIncome", 3, 0.1m, 100m);
            vatInput = CreateNumber("numCostSummaryVat", 3, 0.1m, 100m);
            componentGrid = CreateComponentGrid();

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));

            var header = new Panel { Dock = DockStyle.Fill };
            header.Controls.Add(new Label
            {
                Text = "Tổng hợp kinh phí thi công",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 12f, FontStyle.Bold)
            });
            header.Controls.Add(new Label
            {
                Text = context.PackageId + " v" + context.PackageVersion + " | " +
                    ShortChecksum(context.PackageChecksum),
                Dock = DockStyle.Bottom,
                Height = 20,
                ForeColor = SystemColors.GrayText
            });
            root.Controls.Add(header, 0, 0);

            var inputs = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 4,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            for (int column = 0; column < 4; column++)
                inputs.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            for (int row = 0; row < 4; row++)
                inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 25));
            AddInput(inputs, 0, 0, "Biểu mẫu", templateInput);
            AddInput(inputs, 1, 0, "Địa hình", terrainInput);
            AddInput(inputs, 2, 0, "Diện tích (ha)", areaInput);
            AddInput(inputs, 3, 0, "Loại dự án", projectInput);
            AddInput(inputs, 0, 2, "Loại công trình", constructionInput);
            AddInput(inputs, 1, 2, "Khối lượng hủy nổ", disposalInput);
            AddInput(inputs, 2, 2, "TL (%)", preTaxIncomeInput);
            AddInput(inputs, 3, 2, "VAT (%)", vatInput);
            root.Controls.Add(inputs, 0, 1);
            root.Controls.Add(componentGrid, 0, 2);

            var summary = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                RowCount = 1,
                Margin = new Padding(0, 6, 0, 0)
            };
            for (int column = 0; column < 6; column++)
                summary.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.6667f));
            directValue = AddSummary(summary, 0, "Trực tiếp (T)");
            zValue = AddSummary(summary, 1, "Cơ sở (Z)");
            otherValue = AddSummary(summary, 2, "Chi phí K");
            beforeTaxValue = AddSummary(summary, 3, "Trước VAT (Q)");
            vatValue = AddSummary(summary, 4, "VAT");
            finalValue = AddSummary(summary, 5, "Làm tròn");
            root.Controls.Add(summary, 0, 3);

            statusLabel = new Label
            {
                Name = "lblCostSummaryStatus",
                Text = context.ImportNote,
                Dock = DockStyle.Fill,
                ForeColor = Color.DarkGoldenrod,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
            root.Controls.Add(statusLabel, 0, 4);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0)
            };
            var writeButton = new Button
            {
                Name = "btnWriteCostSummary",
                Text = "Ghi vào THKP-TC",
                AutoSize = true,
                Height = 30
            };
            var previewButton = new Button
            {
                Name = "btnPreviewCostSummary",
                Text = "Tính lại",
                AutoSize = true,
                Height = 30
            };
            writeButton.Click += WriteButton_Click;
            previewButton.Click += (sender, args) => Recalculate();
            buttons.Controls.Add(writeButton);
            buttons.Controls.Add(previewButton);
            root.Controls.Add(buttons, 0, 5);
            Controls.Add(root);

            PopulateOptions();
            PopulateComponents();
            LoadDefaultRequest(context.DefaultRequest);
            templateInput.SelectedIndexChanged += (sender, args) => UpdateTemplateState();
            componentGrid.CurrentCellDirtyStateChanged += (sender, args) =>
            {
                if (componentGrid.IsCurrentCellDirty)
                    componentGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            UpdateTemplateState();
            Recalculate();
        }

        public WorkbookCostSummaryPreview CurrentPreview { get; private set; }

        public void Recalculate()
        {
            try
            {
                CurrentPreview = WorkbookCostSummaryService.Preview(context, BuildRequest());
                ShowResult(CurrentPreview.Result);
                statusLabel.Text = context.ImportNote;
                statusLabel.ForeColor = Color.DarkGoldenrod;
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Preview cost summary", "DT-407", "preview", workbook, "CostSummary");
                CurrentPreview = null;
                statusLabel.Text = ex.Message;
                statusLabel.ForeColor = Color.Firebrick;
            }
        }

        private void WriteButton_Click(object sender, EventArgs e)
        {
            try
            {
                Recalculate();
                if (CurrentPreview == null)
                    return;
                WorkbookCostSummaryWriteResult written = WorkbookCostSummaryWriter.Apply(workbook, CurrentPreview);
                statusLabel.Text = "Đã ghi " + written.WrittenAddress + " trên sheet " + written.WorksheetName + ".";
                statusLabel.ForeColor = Color.DarkGreen;
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Write cost summary", "DT-407", "write", workbook, "CostSummary");
                statusLabel.Text = ex.Message;
                statusLabel.ForeColor = Color.Firebrick;
                MessageBox.Show(ex.Message, "Không ghi được tổng hợp kinh phí", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private CostSummaryCalculationRequest BuildRequest()
        {
            CostComponentSelection selection = CostComponentSelection.None;
            foreach (DataGridViewRow row in componentGrid.Rows)
            {
                if (!Convert.ToBoolean(row.Cells["colSelected"].Value ?? false, CultureInfo.InvariantCulture))
                    continue;
                string code = Convert.ToString(row.Cells["colCode"].Value, CultureInfo.InvariantCulture);
                selection |= (CostComponentSelection)(1 << (code[1] - '1'));
            }
            return new CostSummaryCalculationRequest(
                ((OptionItem<CostSummaryTemplate>)templateInput.SelectedItem).Value,
                context.DefaultRequest.MaterialVnd,
                context.DefaultRequest.LaborVnd,
                context.DefaultRequest.MachineVnd,
                ((OptionItem<string>)terrainInput.SelectedItem).Value,
                areaInput.Value,
                ((OptionItem<CostProjectKind>)projectInput.SelectedItem).Value,
                ((OptionItem<CostConstructionKind>)constructionInput.SelectedItem).Value,
                ((OptionItem<decimal>)disposalInput.SelectedItem).Value,
                preTaxIncomeInput.Value,
                vatInput.Value,
                selection);
        }

        private void ShowResult(CostSummaryCalculationResult result)
        {
            foreach (DataGridViewRow row in componentGrid.Rows)
            {
                string code = Convert.ToString(row.Cells["colCode"].Value, CultureInfo.InvariantCulture);
                CostComponentCalculationResult component = result.Components.FirstOrDefault(item => item.ComponentCode == code);
                row.Cells["colRate"].Value = component?.CalculatedRatePercent?.ToString("0.####", culture) ?? string.Empty;
                row.Cells["colBasis"].Value = code == "K2" ? "T" : "Z";
                row.Cells["colAmount"].Value = component == null ? string.Empty : FormatMoney(component.AppliedAmountVnd);
                row.Cells["colSource"].Value = component == null
                    ? string.Empty
                    : component.Source.DocumentId + " tr." + component.Source.PageFrom.ToString(culture);
            }
            directValue.Text = FormatMoney(result.DirectCost.DirectVnd);
            zValue.Text = FormatMoney(result.ZVnd);
            otherValue.Text = FormatMoney(result.OtherCostTotalVnd);
            beforeTaxValue.Text = FormatMoney(result.BeforeTaxVnd);
            vatValue.Text = FormatMoney(result.VatVnd);
            finalValue.Text = FormatMoney(result.RoundedAfterTaxVnd);
        }

        private void PopulateOptions()
        {
            templateInput.Items.Add(new OptionItem<CostSummaryTemplate>("01 - Điều tra, khảo sát", CostSummaryTemplate.Survey));
            templateInput.Items.Add(new OptionItem<CostSummaryTemplate>("02 - Dự án độc lập vốn Nhà nước", CostSummaryTemplate.IndependentStateFunded));
            templateInput.Items.Add(new OptionItem<CostSummaryTemplate>("03 - Hạng mục vốn Nhà nước", CostSummaryTemplate.ProjectItemStateFunded));
            templateInput.Items.Add(new OptionItem<CostSummaryTemplate>("04 - Nguồn vốn khác", CostSummaryTemplate.OtherFunding));
            foreach (CostTerrainDefinition terrain in context.Catalog.TerrainDefinitions)
                terrainInput.Items.Add(new OptionItem<string>(TerrainLabel(terrain.Terrain), terrain.Terrain));
            projectInput.Items.Add(new OptionItem<CostProjectKind>("Công trình theo tuyến", CostProjectKind.Linear));
            projectInput.Items.Add(new OptionItem<CostProjectKind>("Công trình còn lại", CostProjectKind.Other));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Dân dụng", CostConstructionKind.Civil));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Công nghiệp", CostConstructionKind.Industrial));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Giao thông", CostConstructionKind.Transport));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Nông nghiệp và môi trường", CostConstructionKind.AgricultureAndEnvironment));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Hạ tầng kỹ thuật", CostConstructionKind.TechnicalInfrastructure));
            disposalInput.Items.Add(new OptionItem<decimal>("Dưới 1.000 kg", 999m));
            disposalInput.Items.Add(new OptionItem<decimal>("Trên 1.000 kg", 1001m));
        }

        private void PopulateComponents()
        {
            foreach (string code in new[] { "K1", "K2", "K3", "K4", "K5", "K6" })
                componentGrid.Rows.Add(true, code, string.Empty, string.Empty, string.Empty, string.Empty);
        }

        private void LoadDefaultRequest(CostSummaryCalculationRequest request)
        {
            Select(templateInput, request.Template);
            Select(terrainInput, request.Terrain);
            areaInput.Value = Clamp(areaInput, request.AreaHa);
            Select(projectInput, request.ProjectKind);
            Select(constructionInput, request.ConstructionKind);
            Select(disposalInput, request.DisposalWeightKg < 1000m ? 999m : 1001m);
            preTaxIncomeInput.Value = Clamp(preTaxIncomeInput, request.PreTaxIncomeRatePercent);
            vatInput.Value = Clamp(vatInput, request.VatRatePercent);
        }

        private void UpdateTemplateState()
        {
            bool otherFunding = templateInput.SelectedItem is OptionItem<CostSummaryTemplate> item &&
                item.Value == CostSummaryTemplate.OtherFunding;
            preTaxIncomeInput.Enabled = otherFunding;
            vatInput.Enabled = otherFunding;
        }

        private string FormatMoney(long value)
        {
            return value.ToString("N0", culture);
        }

        private static DataGridView CreateComponentGrid()
        {
            var grid = new DataGridView
            {
                Name = "gridCostSummaryComponents",
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle
            };
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "colSelected", HeaderText = "Dùng", Width = 48 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colCode", HeaderText = "Mã", Width = 48, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colRate", HeaderText = "Tỷ lệ (%)", Width = 78, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colBasis", HeaderText = "Cơ sở", Width = 58, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colAmount", HeaderText = "Giá trị", Width = 130, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colSource", HeaderText = "Căn cứ", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, ReadOnly = true });
            return grid;
        }

        private static ComboBox CreateComboBox(string name)
        {
            return new ComboBox { Name = name, Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
        }

        private static NumericUpDown CreateNumber(
            string name,
            int decimalPlaces,
            decimal increment,
            decimal maximum)
        {
            return new NumericUpDown
            {
                Name = name,
                Dock = DockStyle.Fill,
                DecimalPlaces = decimalPlaces,
                Increment = increment,
                Minimum = 0m,
                Maximum = maximum,
                ThousandsSeparator = true
            };
        }

        private static void AddInput(TableLayoutPanel panel, int column, int row, string title, Control control)
        {
            panel.Controls.Add(new Label
            {
                Text = title,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.BottomLeft,
                AutoEllipsis = true
            }, column, row);
            panel.Controls.Add(control, column, row + 1);
        }

        private static Label AddSummary(TableLayoutPanel panel, int column, string title)
        {
            var container = new Panel { Dock = DockStyle.Fill, Padding = new Padding(4, 3, 6, 3) };
            var titleLabel = new Label { Text = title, Dock = DockStyle.Top, Height = 21, ForeColor = SystemColors.GrayText };
            var value = new Label
            {
                Text = "0",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                AutoEllipsis = true
            };
            container.Controls.Add(value);
            container.Controls.Add(titleLabel);
            panel.Controls.Add(container, column, 0);
            return value;
        }

        private static decimal Clamp(NumericUpDown input, decimal value)
        {
            return Math.Max(input.Minimum, Math.Min(input.Maximum, value));
        }

        private static void Select<T>(ComboBox comboBox, T value)
        {
            for (int index = 0; index < comboBox.Items.Count; index++)
            {
                OptionItem<T> item = comboBox.Items[index] as OptionItem<T>;
                if (item != null && Equals(item.Value, value))
                {
                    comboBox.SelectedIndex = index;
                    return;
                }
            }
            if (comboBox.Items.Count > 0)
                comboBox.SelectedIndex = 0;
        }

        private static string TerrainLabel(string value)
        {
            switch (value)
            {
                case "plain-open": return "Đồng bằng, trống trải";
                case "urban-residential": return "Đô thị, khu dân cư";
                case "midland-or-forest-1": return "Trung du hoặc rừng loại 1";
                case "forest-2": return "Rừng loại 2";
                case "forest-3": return "Rừng loại 3";
                case "forest-4": return "Rừng loại 4";
                case "underwater": return "Dưới nước";
                case "offshore": return "Dưới biển";
                default: return value;
            }
        }

        private static string ShortChecksum(string checksum)
        {
            return string.IsNullOrEmpty(checksum) || checksum.Length <= 12
                ? checksum
                : checksum.Substring(0, 12) + "...";
        }

        private sealed class OptionItem<T>
        {
            public OptionItem(string text, T value)
            {
                Text = text;
                Value = value;
            }

            public string Text { get; }
            public T Value { get; }
            public override string ToString() => Text;
        }
    }
}
