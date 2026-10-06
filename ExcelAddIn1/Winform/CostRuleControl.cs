using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace ExcelAddIn1.Winform
{
    public sealed class CostRuleControl : UserControl
    {
        private readonly WorkbookCostRuleContext context;
        private readonly NumericUpDown materialInput;
        private readonly NumericUpDown laborInput;
        private readonly NumericUpDown machineInput;
        private readonly NumericUpDown areaInput;
        private readonly NumericUpDown disposalInput;
        private readonly NumericUpDown vatInput;
        private readonly ComboBox terrainInput;
        private readonly ComboBox projectInput;
        private readonly ComboBox constructionInput;
        private readonly DataGridView componentGrid;
        private readonly Label directValue;
        private readonly Label commonValue;
        private readonly Label zValue;
        private readonly Label otherValue;
        private readonly Label beforeTaxValue;
        private readonly Label vatValue;
        private readonly Label afterTaxValue;
        private readonly Label statusLabel;

        public CostRuleControl(WorkbookCostRuleContext context)
        {
            this.context = context ?? throw new ArgumentNullException(nameof(context));
            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;

            var inputPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 150,
                Padding = new Padding(10, 8, 10, 4),
                ColumnCount = 6,
                RowCount = 4
            };
            for (int index = 0; index < 6; index++)
                inputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.666f));
            inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            inputPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

            materialInput = CreateMoneyInput("numMaterial", 300000000m);
            laborInput = CreateMoneyInput("numLabor", 200000000m);
            machineInput = CreateMoneyInput("numMachine", 100000000m);
            areaInput = CreateDecimalInput("numArea", 10m, 4, 0.0001m, 100000000m);
            disposalInput = CreateDecimalInput("numDisposal", 500m, 2, 0.01m, 100000000m);
            vatInput = CreateDecimalInput("numVat", 10m, 2, 0.01m, 100m);
            terrainInput = CreateComboBox("cboTerrain");
            projectInput = CreateComboBox("cboProject");
            constructionInput = CreateComboBox("cboConstruction");

            AddInput(inputPanel, 0, "Vật liệu (đ)", materialInput);
            AddInput(inputPanel, 1, "Nhân công (đ)", laborInput);
            AddInput(inputPanel, 2, "Máy thi công (đ)", machineInput);
            AddInput(inputPanel, 3, "Diện tích (ha)", areaInput);
            AddInput(inputPanel, 4, "Khối lượng xử lý (kg)", disposalInput);
            AddInput(inputPanel, 5, "VAT (%)", vatInput);
            AddInput(inputPanel, 0, "Địa hình", terrainInput, 2);
            AddInput(inputPanel, 2, "Loại dự án", projectInput, 2);
            AddInput(inputPanel, 4, "Loại công trình", constructionInput, 2);

            componentGrid = CreateComponentGrid();
            var resultPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 4, 10, 4)
            };
            resultPanel.Controls.Add(componentGrid);

            var summaryPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 104,
                Padding = new Padding(10, 6, 10, 4),
                ColumnCount = 4,
                RowCount = 2
            };
            summaryPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            summaryPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            summaryPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            summaryPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            summaryPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            summaryPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            directValue = AddSummary(summaryPanel, 0, 0, "Trực tiếp (T)");
            commonValue = AddSummary(summaryPanel, 1, 0, "Chi phí chung (C)");
            zValue = AddSummary(summaryPanel, 2, 0, "T + C (Z)");
            otherValue = AddSummary(summaryPanel, 3, 0, "Tổng K1-K6");
            beforeTaxValue = AddSummary(summaryPanel, 0, 1, "Trước thuế");
            vatValue = AddSummary(summaryPanel, 1, 1, "Thuế GTGT");
            afterTaxValue = AddSummary(summaryPanel, 2, 1, "Sau thuế");

            var commandPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                Padding = new Padding(10, 4, 10, 4)
            };
            var calculateButton = new Button
            {
                Name = "btnCalculateCost",
                Text = "Tính chi phí",
                Dock = DockStyle.Right,
                Width = 118
            };
            calculateButton.Click += (sender, args) => Recalculate();
            statusLabel = new Label
            {
                Name = "lblCostStatus",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
            commandPanel.Controls.Add(statusLabel);
            commandPanel.Controls.Add(calculateButton);

            Controls.Add(resultPanel);
            Controls.Add(inputPanel);
            Controls.Add(summaryPanel);
            Controls.Add(commandPanel);

            PopulateOptions();
            PopulateComponents();
            Recalculate();
        }

        public CostRuleEngineResult CurrentResult { get; private set; }

        public void Recalculate()
        {
            try
            {
                CostRuleCalculationRequest request = BuildRequest();
                CurrentResult = CostRuleEngine.Calculate(context.Catalog, request);
                ShowResult(CurrentResult);
                statusLabel.ForeColor = SystemColors.GrayText;
                statusLabel.Text = "Package: " + context.PackageId + " v" + context.PackageVersion;
            }
            catch (Exception ex)
            {
                CurrentResult = null;
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
            }
        }

        private CostRuleCalculationRequest BuildRequest()
        {
            CostComponentSelection selection = CostComponentSelection.None;
            var overrides = new List<CostRuleOverride>();
            foreach (DataGridViewRow row in componentGrid.Rows)
            {
                string code = Convert.ToString(row.Cells["colCode"].Value, CultureInfo.InvariantCulture);
                bool selected = Convert.ToBoolean(row.Cells["colSelected"].Value ?? false, CultureInfo.InvariantCulture);
                if (!selected)
                    continue;
                selection |= (CostComponentSelection)Enum.Parse(typeof(CostComponentSelection), code, false);
                bool overridden = Convert.ToBoolean(row.Cells["colOverride"].Value ?? false, CultureInfo.InvariantCulture);
                if (overridden)
                {
                    string amountText = Convert.ToString(row.Cells["colApplied"].Value, CultureInfo.CurrentCulture);
                    long amount;
                    if (!long.TryParse(amountText, NumberStyles.Number, CultureInfo.CurrentCulture, out amount))
                        throw new ArgumentException("Giá trị ghi đè " + code + " không hợp lệ.");
                    string reason = Convert.ToString(row.Cells["colReason"].Value, CultureInfo.CurrentCulture);
                    overrides.Add(new CostRuleOverride(code, amount, reason));
                }
            }

            return new CostRuleCalculationRequest(
                materialInput.Value,
                laborInput.Value,
                machineInput.Value,
                ((OptionItem<string>)terrainInput.SelectedItem).Value,
                areaInput.Value,
                ((OptionItem<CostProjectKind>)projectInput.SelectedItem).Value,
                ((OptionItem<CostConstructionKind>)constructionInput.SelectedItem).Value,
                disposalInput.Value,
                vatInput.Value,
                selection,
                overrides);
        }

        private void ShowResult(CostRuleEngineResult result)
        {
            foreach (DataGridViewRow row in componentGrid.Rows)
            {
                string code = Convert.ToString(row.Cells["colCode"].Value, CultureInfo.InvariantCulture);
                CostComponentCalculationResult component = result.Components.FirstOrDefault(item => item.ComponentCode == code);
                if (component == null)
                {
                    row.Cells["colRate"].Value = string.Empty;
                    row.Cells["colCalculated"].Value = string.Empty;
                    row.Cells["colSource"].Value = string.Empty;
                    continue;
                }
                row.Cells["colRate"].Value = component.CalculatedRatePercent?.ToString("0.####", CultureInfo.CurrentCulture);
                row.Cells["colCalculated"].Value = FormatMoney(component.CalculatedAmountVnd);
                row.Cells["colApplied"].Value = FormatMoney(component.AppliedAmountVnd);
                row.Cells["colSource"].Value = FormatSource(component);
            }
            directValue.Text = FormatMoney(result.DirectCost.DirectVnd);
            commonValue.Text = FormatMoney(result.DirectCost.CommonVnd);
            zValue.Text = FormatMoney(result.DirectCost.ZVnd);
            otherValue.Text = FormatMoney(result.Summary.OtherCostTotalVnd);
            beforeTaxValue.Text = FormatMoney(result.Summary.BeforeTaxVnd);
            vatValue.Text = FormatMoney(result.Summary.VatVnd);
            afterTaxValue.Text = FormatMoney(result.Summary.AfterTaxVnd);
        }

        private void PopulateOptions()
        {
            foreach (CostTerrainDefinition terrain in context.Catalog.TerrainDefinitions)
                terrainInput.Items.Add(new OptionItem<string>(TerrainLabel(terrain.Terrain), terrain.Terrain));
            terrainInput.SelectedIndex = 0;
            projectInput.Items.Add(new OptionItem<CostProjectKind>("Dự án theo tuyến", CostProjectKind.Linear));
            projectInput.Items.Add(new OptionItem<CostProjectKind>("Dự án khác", CostProjectKind.Other));
            projectInput.SelectedIndex = 0;
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Dân dụng", CostConstructionKind.Civil));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Công nghiệp", CostConstructionKind.Industrial));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Giao thông", CostConstructionKind.Transport));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Nông nghiệp và môi trường", CostConstructionKind.AgricultureAndEnvironment));
            constructionInput.Items.Add(new OptionItem<CostConstructionKind>("Hạ tầng kỹ thuật", CostConstructionKind.TechnicalInfrastructure));
            constructionInput.SelectedIndex = 0;
        }

        private void PopulateComponents()
        {
            foreach (string code in new[] { "K1", "K2", "K3", "K4", "K5", "K6" })
            {
                int index = componentGrid.Rows.Add(
                    true, code, string.Empty, string.Empty, false, string.Empty, string.Empty, string.Empty);
                DataGridViewRow row = componentGrid.Rows[index];
                row.Cells["colApplied"].ReadOnly = true;
                row.Cells["colReason"].ReadOnly = true;
                row.Cells["colApplied"].Style.BackColor = SystemColors.Control;
                row.Cells["colReason"].Style.BackColor = SystemColors.Control;
            }
        }

        private static DataGridView CreateComponentGrid()
        {
            var grid = new DataGridView
            {
                Name = "gridCostComponents",
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
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "colSelected", HeaderText = "Dùng", Width = 46 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colCode", HeaderText = "Mã", Width = 44, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colRate", HeaderText = "Tỷ lệ (%)", Width = 70, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colCalculated", HeaderText = "Theo quy định", Width = 112, ReadOnly = true });
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "colOverride", HeaderText = "Ghi đè", Width = 56 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colApplied", HeaderText = "Giá trị áp dụng", Width = 116 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colReason", HeaderText = "Lý do ghi đè", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "colSource", HeaderText = "Căn cứ", Width = 175, ReadOnly = true });
            grid.CellValueChanged += (sender, args) =>
            {
                if (args.RowIndex < 0)
                    return;
                DataGridViewRow row = grid.Rows[args.RowIndex];
                bool enabled = Convert.ToBoolean(row.Cells["colOverride"].Value ?? false, CultureInfo.InvariantCulture);
                row.Cells["colApplied"].ReadOnly = !enabled;
                row.Cells["colReason"].ReadOnly = !enabled;
                row.Cells["colApplied"].Style.BackColor = enabled ? SystemColors.Window : SystemColors.Control;
                row.Cells["colReason"].Style.BackColor = enabled ? SystemColors.Window : SystemColors.Control;
            };
            grid.CurrentCellDirtyStateChanged += (sender, args) =>
            {
                if (grid.IsCurrentCellDirty)
                    grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            return grid;
        }

        private static NumericUpDown CreateMoneyInput(string name, decimal value)
        {
            return CreateDecimalInput(name, value, 0, 1000m, 1000000000000000m);
        }

        private static NumericUpDown CreateDecimalInput(
            string name,
            decimal value,
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
                Maximum = maximum,
                Minimum = 0m,
                ThousandsSeparator = true,
                Value = value
            };
        }

        private static ComboBox CreateComboBox(string name)
        {
            return new ComboBox
            {
                Name = name,
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
        }

        private static void AddInput(
            TableLayoutPanel panel,
            int column,
            string title,
            Control control,
            int columnSpan = 1)
        {
            int row = columnSpan == 1 ? 0 : 2;
            panel.Controls.Add(new Label
            {
                Text = title,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            }, column, row);
            panel.Controls.Add(control, column, row + 1);
            if (columnSpan > 1)
            {
                panel.SetColumnSpan(panel.GetControlFromPosition(column, row), columnSpan);
                panel.SetColumnSpan(control, columnSpan);
            }
        }

        private static Label AddSummary(TableLayoutPanel panel, int column, int row, string title)
        {
            var container = new Panel { Dock = DockStyle.Fill, Padding = new Padding(4, 2, 8, 2) };
            var titleLabel = new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                Height = 20,
                ForeColor = SystemColors.GrayText
            };
            var valueLabel = new Label
            {
                Text = "0",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold)
            };
            container.Controls.Add(valueLabel);
            container.Controls.Add(titleLabel);
            panel.Controls.Add(container, column, row);
            return valueLabel;
        }

        private static string FormatMoney(long value)
        {
            return value.ToString("N0", CultureInfo.CurrentCulture);
        }

        private static string FormatSource(CostComponentCalculationResult component)
        {
            string source = component.Source.DocumentId + " tr." +
                component.Source.PageFrom.ToString(CultureInfo.CurrentCulture);
            if (component.Source.PageTo != component.Source.PageFrom)
                source += "-" + component.Source.PageTo.ToString(CultureInfo.CurrentCulture);
            if (component.ExternalBasis.Length > 0)
                source += "; " + component.ExternalBasis;
            return source;
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
