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
    public sealed class UnitRateControl : UserControl
    {
        private readonly WorkbookUnitRateContext context;
        private readonly CultureInfo numberCulture;
        private readonly RadioButton landMode;
        private readonly RadioButton waterMode;
        private readonly ComboBox normInput;
        private readonly ComboBox variantInput;
        private readonly CheckBox applyConditionsInput;
        private readonly CheckedListBox conditionList;
        private readonly DataGridView bindingGrid;
        private readonly DataGridView resultGrid;
        private readonly Label identityLabel;
        private readonly Label sourceLabel;
        private readonly Label materialTotalLabel;
        private readonly Label laborTotalLabel;
        private readonly Label machineTotalLabel;
        private readonly Label grandTotalLabel;
        private readonly Label statusLabel;
        private bool loading;

        public UnitRateControl(WorkbookUnitRateContext context, CultureInfo numberCulture)
        {
            this.context = context ?? throw new ArgumentNullException(nameof(context));
            this.numberCulture = numberCulture ?? CultureInfo.CurrentCulture;
            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;

            var topPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 235,
                Padding = new Padding(10, 8, 10, 4),
                ColumnCount = 3,
                RowCount = 1
            };
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 36));
            var selectorColumn = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 8, 0) };
            var conditionColumn = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 8, 0) };
            var bindingColumn = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            var resultPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 4, 10, 8) };

            identityLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 42,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.GrayText,
                Text = context.PackageId + "@" + context.PackageVersion + Environment.NewLine +
                    context.PriceProfile.ProfileId + "@" + context.PriceProfile.DataVersion
            };

            var modePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 34,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            landMode = new RadioButton
            {
                Name = "radUnitRateLand",
                Text = "Trên cạn",
                AutoSize = true,
                Checked = true,
                Margin = new Padding(3, 7, 16, 3)
            };
            waterMode = new RadioButton
            {
                Name = "radUnitRateWater",
                Text = "Dưới nước",
                AutoSize = true,
                Margin = new Padding(3, 7, 3, 3)
            };
            modePanel.Controls.Add(landMode);
            modePanel.Controls.Add(waterMode);

            var selectorPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 102,
                ColumnCount = 1,
                RowCount = 4
            };
            selectorPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            selectorPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 29));
            selectorPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            selectorPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 29));
            selectorPanel.Controls.Add(CreateLabel("Định mức"), 0, 0);
            normInput = CreateComboBox("cboUnitRateNorm");
            selectorPanel.Controls.Add(normInput, 0, 1);
            selectorPanel.Controls.Add(CreateLabel("Biến thể"), 0, 2);
            variantInput = CreateComboBox("cboUnitRateVariant");
            selectorPanel.Controls.Add(variantInput, 0, 3);

            applyConditionsInput = new CheckBox
            {
                Name = "chkApplyUnitRateConditions",
                Text = "Áp dụng điều kiện / hệ số",
                Dock = DockStyle.Top,
                Height = 30,
                Checked = false,
                Padding = new Padding(0, 5, 0, 0)
            };
            conditionList = new CheckedListBox
            {
                Name = "clbUnitRateConditions",
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                IntegralHeight = false,
                Enabled = false
            };
            var bindingLabel = CreateLabel("Lựa chọn nguồn lực");
            bindingLabel.Dock = DockStyle.Top;
            bindingLabel.Height = 24;
            bindingGrid = CreateBindingGrid();
            bindingGrid.Dock = DockStyle.Fill;

            var calculateButton = new Button
            {
                Name = "btnCalculateUnitRate",
                Text = "Tính đơn giá",
                Dock = DockStyle.Bottom,
                Height = 34
            };
            calculateButton.Click += (sender, args) => CalculateCurrent();

            selectorColumn.Controls.Add(selectorPanel);
            selectorColumn.Controls.Add(modePanel);
            selectorColumn.Controls.Add(identityLabel);
            selectorColumn.Controls.Add(calculateButton);
            conditionColumn.Controls.Add(conditionList);
            conditionColumn.Controls.Add(applyConditionsInput);
            bindingColumn.Controls.Add(bindingGrid);
            bindingColumn.Controls.Add(bindingLabel);
            topPanel.Controls.Add(selectorColumn, 0, 0);
            topPanel.Controls.Add(conditionColumn, 1, 0);
            topPanel.Controls.Add(bindingColumn, 2, 0);

            resultGrid = CreateResultGrid();
            var totals = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 58,
                ColumnCount = 4,
                RowCount = 2,
                Padding = new Padding(0, 3, 0, 0)
            };
            for (int index = 0; index < 4; index++)
                totals.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            totals.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            totals.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            totals.Controls.Add(CreateCenteredLabel("Vật liệu"), 0, 0);
            totals.Controls.Add(CreateCenteredLabel("Nhân công"), 1, 0);
            totals.Controls.Add(CreateCenteredLabel("Máy"), 2, 0);
            totals.Controls.Add(CreateCenteredLabel("Tổng"), 3, 0);
            materialTotalLabel = CreateTotalLabel(false);
            laborTotalLabel = CreateTotalLabel(false);
            machineTotalLabel = CreateTotalLabel(false);
            grandTotalLabel = CreateTotalLabel(true);
            totals.Controls.Add(materialTotalLabel, 0, 1);
            totals.Controls.Add(laborTotalLabel, 1, 1);
            totals.Controls.Add(machineTotalLabel, 2, 1);
            totals.Controls.Add(grandTotalLabel, 3, 1);

            statusLabel = new Label
            {
                Name = "lblUnitRateStatus",
                Dock = DockStyle.Bottom,
                Height = 38,
                AutoEllipsis = true,
                Padding = new Padding(4, 5, 4, 3),
                ForeColor = SystemColors.GrayText
            };
            sourceLabel = new Label
            {
                Name = "lblUnitRateSource",
                Dock = DockStyle.Top,
                Height = 42,
                AutoEllipsis = true,
                Padding = new Padding(4, 3, 4, 3),
                ForeColor = SystemColors.GrayText
            };
            resultPanel.Controls.Add(resultGrid);
            resultPanel.Controls.Add(sourceLabel);
            resultPanel.Controls.Add(statusLabel);
            resultPanel.Controls.Add(totals);
            Controls.Add(resultPanel);
            Controls.Add(topPanel);

            landMode.CheckedChanged += (sender, args) =>
            {
                if (landMode.Checked && !loading)
                    LoadDefinitions();
            };
            waterMode.CheckedChanged += (sender, args) =>
            {
                if (waterMode.Checked && !loading)
                    LoadDefinitions();
            };
            normInput.SelectedIndexChanged += (sender, args) =>
            {
                if (!loading)
                    LoadDefinitionDetails();
            };
            applyConditionsInput.CheckedChanged += (sender, args) =>
                conditionList.Enabled = applyConditionsInput.Checked;

            LoadDefinitions();
        }

        public UnitRateCalculationResult CurrentResult { get; private set; }

        public void SetWaterMode(bool water)
        {
            if (water)
                waterMode.Checked = true;
            else
                landMode.Checked = true;
        }

        public bool SelectDefinition(string normKey, string variantCode)
        {
            string wanted = (normKey ?? string.Empty).Trim();
            for (int index = 0; index < normInput.Items.Count; index++)
            {
                NormOption option = (NormOption)normInput.Items[index];
                if (!string.Equals(option.Definition.Key, wanted, StringComparison.Ordinal))
                    continue;
                normInput.SelectedIndex = index;
                for (int variantIndex = 0; variantIndex < variantInput.Items.Count; variantIndex++)
                {
                    if (string.Equals(
                        ((ValueOption)variantInput.Items[variantIndex]).Value,
                        variantCode,
                        StringComparison.Ordinal))
                    {
                        variantInput.SelectedIndex = variantIndex;
                        return true;
                    }
                }
                return false;
            }
            return false;
        }

        public bool SetCondition(string conditionCode, bool enabled)
        {
            applyConditionsInput.Checked = true;
            for (int index = 0; index < conditionList.Items.Count; index++)
            {
                ConditionOption option = (ConditionOption)conditionList.Items[index];
                if (string.Equals(option.Code, conditionCode, StringComparison.Ordinal))
                {
                    conditionList.SetItemChecked(index, enabled);
                    return true;
                }
            }
            return false;
        }

        public bool SetBinding(string resourceCode, string priceCode, string reason)
        {
            foreach (DataGridViewRow row in bindingGrid.Rows)
            {
                if (string.Equals(
                    Convert.ToString(row.Cells["colBindingResource"].Value, CultureInfo.InvariantCulture),
                    resourceCode,
                    StringComparison.OrdinalIgnoreCase))
                {
                    row.Cells["colBindingPrice"].Value = priceCode;
                    row.Cells["colBindingReason"].Value = reason;
                    return true;
                }
            }
            return false;
        }

        public UnitRateCalculationResult CalculateCurrent()
        {
            resultGrid.Rows.Clear();
            CurrentResult = null;
            try
            {
                UnitRateCalculationRequest request = BuildRequest();
                UnitRateValidationResult validation = UnitRateCalculator.Validate(request);
                if (!validation.IsValid)
                {
                    ShowIssues(validation.Issues);
                    ClearTotals();
                    return null;
                }
                CurrentResult = UnitRateCalculator.Calculate(request);
                ShowResult(CurrentResult);
                return CurrentResult;
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Calculate unit rate");
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
                ClearTotals();
                return null;
            }
        }

        private void LoadDefinitions()
        {
            loading = true;
            try
            {
                normInput.Items.Clear();
                IEnumerable<NormDefinition> definitions = context.NormCatalog.Definitions
                    .Where(definition => IsInMode(definition.Key, landMode.Checked))
                    .OrderBy(definition => definition.Key, StringComparer.Ordinal);
                foreach (NormDefinition definition in definitions)
                    normInput.Items.Add(new NormOption(definition));
                if (normInput.Items.Count > 0)
                    normInput.SelectedIndex = 0;
            }
            finally
            {
                loading = false;
            }
            LoadDefinitionDetails();
        }

        private void LoadDefinitionDetails()
        {
            NormDefinition definition = SelectedDefinition;
            loading = true;
            try
            {
                variantInput.Items.Clear();
                conditionList.Items.Clear();
                bindingGrid.Rows.Clear();
                if (definition == null)
                    return;
                foreach (string variant in definition.Variants)
                    variantInput.Items.Add(new ValueOption(variant, DescribeVariant(variant)));
                if (variantInput.Items.Count > 0)
                    variantInput.SelectedIndex = 0;
                foreach (ConditionOption condition in GetConditions(definition))
                    conditionList.Items.Add(condition, false);
                foreach (NormResourceRate rate in definition.Rates.Where(
                    item => UnitRateCalculator.IsLogicalResource(item.ResourceCode)))
                {
                    bindingGrid.Rows.Add(rate.ResourceCode, null, string.Empty);
                }
                sourceLabel.Text = FormatSource(definition.Source);
                statusLabel.ForeColor = SystemColors.GrayText;
                statusLabel.Text = "Chọn biến thể và điều kiện để tính đơn giá.";
                resultGrid.Rows.Clear();
                CurrentResult = null;
                ClearTotals();
            }
            finally
            {
                loading = false;
            }
        }

        private UnitRateCalculationRequest BuildRequest()
        {
            NormDefinition definition = SelectedDefinition;
            ValueOption variant = variantInput.SelectedItem as ValueOption;
            if (definition == null || variant == null)
                throw new InvalidOperationException("Chưa chọn định mức và biến thể.");
            string[] conditions = applyConditionsInput.Checked
                ? conditionList.CheckedItems.Cast<ConditionOption>().Select(item => item.Code).ToArray()
                : new string[0];
            var bindings = new List<UnitRateResourceBinding>();
            foreach (DataGridViewRow row in bindingGrid.Rows)
            {
                string resource = CellText(row, "colBindingResource");
                string price = CellText(row, "colBindingPrice");
                string reason = CellText(row, "colBindingReason");
                if (price.Length > 0 || reason.Length > 0)
                    bindings.Add(new UnitRateResourceBinding(resource, price, reason));
            }
            return new UnitRateCalculationRequest(
                definition,
                variant.Value,
                context.PriceProfile,
                conditions,
                bindings);
        }

        private void ShowResult(UnitRateCalculationResult result)
        {
            foreach (UnitRateResourceAmount resource in result.Resources)
            {
                int rowIndex = resultGrid.Rows.Add(
                    KindName(resource.Kind),
                    resource.ResourceCode,
                    resource.PriceCode,
                    resource.DisplayName,
                    resource.Unit,
                    resource.Quantity,
                    resource.IsPercentage ? (object)null : resource.UnitPriceVnd,
                    resource.AmountVnd);
                if (resource.BindingReason.Length > 0)
                {
                    foreach (DataGridViewCell cell in resultGrid.Rows[rowIndex].Cells)
                        cell.ToolTipText = resource.BindingReason;
                }
            }
            materialTotalLabel.Text = Money(result.MaterialAmountVnd);
            laborTotalLabel.Text = Money(result.LaborAmountVnd);
            machineTotalLabel.Text = Money(result.MachineAmountVnd);
            grandTotalLabel.Text = Money(result.TotalAmountVnd);
            statusLabel.ForeColor = Color.DarkGreen;
            statusLabel.Text = result.Resources.Count + " hao phí; " + result.RoundingRule + ".";
        }

        private void ShowIssues(IEnumerable<UnitRateValidationIssue> issues)
        {
            UnitRateValidationIssue[] values = issues.ToArray();
            statusLabel.ForeColor = Color.Firebrick;
            statusLabel.Text = string.Join(" | ", values.Take(3).Select(item => item.Message)) +
                (values.Length > 3 ? " | ..." : string.Empty);
        }

        private void ClearTotals()
        {
            materialTotalLabel.Text = "0";
            laborTotalLabel.Text = "0";
            machineTotalLabel.Text = "0";
            grandTotalLabel.Text = "0";
        }

        private IEnumerable<ConditionOption> GetConditions(NormDefinition definition)
        {
            var options = new Dictionary<string, ConditionOption>(StringComparer.Ordinal);
            foreach (NormAdjustment adjustment in definition.Adjustments)
            {
                string detail = adjustment.Operation == NormAdjustmentOperation.Factor
                    ? "x" + adjustment.Value.ToString("0.##", numberCulture) + " " +
                        string.Join("+", adjustment.TargetKinds.Select(ShortKindName))
                    : "+" + adjustment.Value.ToString("0.###", numberCulture) + " " +
                        adjustment.Unit + " " + adjustment.ResourceCode;
                options[adjustment.Condition] = new ConditionOption(
                    adjustment.Condition,
                    DescribeCondition(adjustment.Condition) + " (" + detail + ")");
            }
            foreach (NormConstraint constraint in definition.Constraints.Where(
                item => item.Operation == "requires-condition"))
            {
                if (!options.ContainsKey(constraint.Argument2))
                {
                    options.Add(constraint.Argument2, new ConditionOption(
                        constraint.Argument2,
                        DescribeCondition(constraint.Argument2) + " (bắt buộc theo biến thể)"));
                }
            }
            return options.Values.OrderBy(item => item.Code, StringComparer.Ordinal);
        }

        private DataGridView CreateBindingGrid()
        {
            var grid = BaseGrid("dgvUnitRateBindings");
            grid.AllowUserToAddRows = false;
            grid.Height = 132;
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colBindingResource",
                HeaderText = "Mã logic",
                Width = 88,
                ReadOnly = true
            });
            var priceColumn = new DataGridViewComboBoxColumn
            {
                Name = "colBindingPrice",
                HeaderText = "Mã giá máy",
                Width = 118,
                FlatStyle = FlatStyle.Flat,
                DisplayMember = "Display",
                ValueMember = "Code",
                DataSource = context.PriceProfile.Entries
                    .Where(item => item.Kind == PriceResourceKind.MachineShift)
                    .OrderBy(item => item.Code, StringComparer.OrdinalIgnoreCase)
                    .Select(item => new PriceOption(item.Code, item.Code + " - " + item.DisplayName))
                    .ToList()
            };
            grid.Columns.Add(priceColumn);
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colBindingReason",
                HeaderText = "Lý do / phương án",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 80
            });
            return grid;
        }

        private DataGridView CreateResultGrid()
        {
            var grid = BaseGrid("dgvUnitRateResult");
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.Columns.Add(TextColumn("colRateKind", "Loại", 48));
            grid.Columns.Add(TextColumn("colRateResource", "Mã hao phí", 108));
            grid.Columns.Add(TextColumn("colRatePriceCode", "Mã giá", 120));
            grid.Columns.Add(TextColumn("colRateName", "Tên nguồn lực", 160));
            grid.Columns.Add(TextColumn("colRateUnit", "ĐVT", 56));
            grid.Columns.Add(NumberColumn("colRateQuantity", "Hao phí", 72, "0.######"));
            grid.Columns.Add(NumberColumn("colRateUnitPrice", "Đơn giá", 98, "#,##0.##"));
            grid.Columns.Add(NumberColumn("colRateAmount", "Thành tiền", 112, "#,##0.##"));
            return grid;
        }

        private DataGridView BaseGrid(string name)
        {
            return new DataGridView
            {
                Name = name,
                Dock = DockStyle.Fill,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                DefaultCellStyle = { FormatProvider = numberCulture }
            };
        }

        private NormDefinition SelectedDefinition => (normInput.SelectedItem as NormOption)?.Definition;

        private static bool IsInMode(string key, bool land)
        {
            if (key.StartsWith("NORM-000.", StringComparison.Ordinal) ||
                key.StartsWith("NORM-010.", StringComparison.Ordinal))
                return true;
            return land
                ? key.StartsWith("NORM-020.", StringComparison.Ordinal)
                : key.StartsWith("NORM-030.", StringComparison.Ordinal) ||
                    key.StartsWith("NORM-040.", StringComparison.Ordinal);
        }

        private static string CellText(DataGridViewRow row, string column)
        {
            return Convert.ToString(row.Cells[column].Value, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
        }

        private string Money(decimal value)
        {
            return value.ToString("#,##0.##", numberCulture);
        }

        private static string KindName(NormResourceKind kind)
        {
            switch (kind)
            {
                case NormResourceKind.Material: return "VL";
                case NormResourceKind.Labor: return "NC";
                case NormResourceKind.Machine: return "M";
                default: return kind.ToString();
            }
        }

        private static string ShortKindName(NormResourceKind kind)
        {
            return KindName(kind);
        }

        private static string DescribeVariant(string code)
        {
            return code.Replace("-", " ");
        }

        private static string DescribeCondition(string code)
        {
            switch (code)
            {
                case "slope-gt-25deg": return "Độ dốc lớn hơn 25 độ";
                case "uxo-signal": return "Tín hiệu bom mìn, vật nổ";
                case "water-excavation": return "Đào đất có nước";
                case "current-gt-0-le-0.5": return "Lưu tốc trên 0 đến 0,5 m/s";
                case "current-gt-0.5-le-1": return "Lưu tốc trên 0,5 đến 1 m/s";
                case "current-gt-1-le-2": return "Lưu tốc trên 1 đến 2 m/s";
                case "forestry-salt-independent-or-owner-request":
                    return "Lâm nghiệp, đồng muối, dự án độc lập hoặc chủ đầu tư yêu cầu";
                default: return code;
            }
        }

        private static string FormatSource(RegulationSourceLocator source)
        {
            string pages = source.PageFrom == source.PageTo
                ? source.PageFrom.ToString(CultureInfo.InvariantCulture)
                : source.PageFrom.ToString(CultureInfo.InvariantCulture) + "-" +
                    source.PageTo.ToString(CultureInfo.InvariantCulture);
            return source.DocumentId + ", trang " + pages + "; " + source.Section;
        }

        private static ComboBox CreateComboBox(string name)
        {
            return new ComboBox
            {
                Name = name,
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                IntegralHeight = false,
                MaxDropDownItems = 14
            };
        }

        private static Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.BottomLeft,
                AutoEllipsis = true
            };
        }

        private static Label CreateCenteredLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
        }

        private static Label CreateTotalLabel(bool emphasized)
        {
            return new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = emphasized
                    ? new Font(SystemFonts.MessageBoxFont, FontStyle.Bold)
                    : SystemFonts.MessageBoxFont,
                ForeColor = emphasized ? Color.DarkGreen : SystemColors.ControlText,
                Text = "0"
            };
        }

        private static DataGridViewTextBoxColumn TextColumn(string name, string title, int width)
        {
            return new DataGridViewTextBoxColumn { Name = name, HeaderText = title, Width = width };
        }

        private static DataGridViewTextBoxColumn NumberColumn(
            string name,
            string title,
            int width,
            string format)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = title,
                Width = width,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleRight, Format = format }
            };
        }

        private sealed class NormOption
        {
            internal NormOption(NormDefinition definition) { Definition = definition; }
            internal NormDefinition Definition { get; }
            public override string ToString() { return Definition.Key + " - " + Definition.Title; }
        }

        private sealed class ValueOption
        {
            internal ValueOption(string value, string display) { Value = value; Display = display; }
            internal string Value { get; }
            internal string Display { get; }
            public override string ToString() { return Display; }
        }

        private sealed class ConditionOption
        {
            internal ConditionOption(string code, string display) { Code = code; Display = display; }
            internal string Code { get; }
            internal string Display { get; }
            public override string ToString() { return Display; }
        }

        private sealed class PriceOption
        {
            internal PriceOption(string code, string display) { Code = code; Display = display; }
            public string Code { get; }
            public string Display { get; }
        }
    }
}
