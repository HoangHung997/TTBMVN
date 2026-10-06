using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class EstimateAppendixControl : UserControl
    {
        private readonly Excel.Workbook workbook;
        private readonly CultureInfo numberCulture;
        private readonly NormCatalog normCatalog;
        private readonly DataGridView lineGrid;
        private readonly ComboBox normInput;
        private readonly ComboBox variantInput;
        private readonly ComboBox environmentInput;
        private readonly ComboBox audienceInput;
        private readonly CheckedListBox conditionInput;
        private readonly DataGridView bindingGrid;
        private readonly Label sourceLabel;
        private readonly Label statusLabel;
        private readonly Label usageLabel;
        private readonly Label totalsLabel;
        private EstimateWorkspace currentWorkspace;
        private WorkbookEstimateWorkspacePreview workspacePreview;
        private bool loadingDetails;

        public EstimateAppendixControl(Excel.Workbook workbook, CultureInfo numberCulture)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.numberCulture = numberCulture ?? CultureInfo.CurrentCulture;
            normCatalog = WorkbookEstimateCalculationService.LoadContext(workbook).NormCatalog;
            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;

            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 86,
                Padding = new Padding(10, 8, 10, 5),
                ColumnCount = 1,
                RowCount = 2
            };
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            sourceLabel = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.GrayText,
                Text = "Chưa quét vùng Phụ lục dự toán từ Excel."
            };
            header.Controls.Add(sourceLabel, 0, 0);
            var commands = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            Button scanButton = Button("btnScanEstimateRange", "Quét vùng Excel", 118);
            Button saveButton = Button("btnSaveEstimateWorkspace", "Lưu bố trí", 94);
            Button previewButton = Button("btnPreviewEstimateWorkspace", "Tính thử", 86);
            Button generateButton = Button("btnGenerateEstimateSheets", "Sinh đơn giá", 105);
            scanButton.Click += (sender, args) => Execute("Quét Phụ lục DT", ScanRange);
            saveButton.Click += (sender, args) => Execute("Lưu Phụ lục DT", () => SaveWorkspace(true));
            previewButton.Click += (sender, args) => Execute("Tính thử Phụ lục DT", RefreshPreview);
            generateButton.Click += (sender, args) => Execute("Sinh bảng đơn giá", GenerateSheets);
            commands.Controls.Add(scanButton);
            commands.Controls.Add(saveButton);
            commands.Controls.Add(previewButton);
            commands.Controls.Add(generateButton);
            totalsLabel = new Label
            {
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(18, 8, 0, 0),
                ForeColor = SystemColors.GrayText
            };
            commands.Controls.Add(totalsLabel);
            header.Controls.Add(commands, 0, 1);

            lineGrid = CreateGrid();
            lineGrid.SelectionChanged += (sender, args) => LoadSelectedDetails();

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 330,
                Panel1MinSize = 180,
                Panel2MinSize = 150
            };
            split.Panel1.Padding = new Padding(10, 0, 10, 3);
            split.Panel1.Controls.Add(lineGrid);

            var detail = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 6, 10, 6),
                ColumnCount = 8,
                RowCount = 3
            };
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 36));
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 76));
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84));
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));
            detail.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            detail.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            detail.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            normInput = DropDown("cboEstimateNorm");
            foreach (NormDefinition definition in normCatalog.Definitions)
                normInput.Items.Add(new NormOption(definition));
            normInput.SelectedIndexChanged += (sender, args) => LoadVariantsAndConditions(null);
            variantInput = DropDown("cboEstimateVariant");
            environmentInput = DropDown("cboEstimateEnvironment");
            environmentInput.Items.Add(new Option<EstimateWorkEnvironment>("Cạn", EstimateWorkEnvironment.Land));
            environmentInput.Items.Add(new Option<EstimateWorkEnvironment>("Nước", EstimateWorkEnvironment.Water));
            audienceInput = DropDown("cboEstimateAudience");
            audienceInput.Items.Add(new Option<MachineRateAudience>(
                "Hưởng lương NSNN", MachineRateAudience.StateBudgetSalary));
            audienceInput.Items.Add(new Option<MachineRateAudience>(
                "Không hưởng lương NSNN", MachineRateAudience.NonStateSalary));

            detail.Controls.Add(DetailLabel("Định mức"), 0, 0);
            detail.Controls.Add(normInput, 1, 0);
            detail.Controls.Add(DetailLabel("Mã chi tiết"), 2, 0);
            detail.Controls.Add(variantInput, 3, 0);
            detail.Controls.Add(DetailLabel("Môi trường"), 4, 0);
            detail.Controls.Add(environmentInput, 5, 0);
            detail.Controls.Add(DetailLabel("Đối tượng"), 6, 0);
            detail.Controls.Add(audienceInput, 7, 0);

            conditionInput = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                IntegralHeight = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            detail.Controls.Add(DetailLabel("Điều kiện/hệ số"), 0, 1);
            detail.Controls.Add(conditionInput, 1, 1);
            detail.SetColumnSpan(conditionInput, 2);
            detail.Controls.Add(DetailLabel("Thiết bị logic"), 3, 1);
            bindingGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false
            };
            bindingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colBindingResource",
                HeaderText = "Mã hao phí",
                Width = 95,
                ReadOnly = true
            });
            bindingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colBindingPrice",
                HeaderText = "Mã giá",
                Width = 105
            });
            bindingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colBindingReason",
                HeaderText = "Lý do chọn",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 100
            });
            detail.Controls.Add(bindingGrid, 4, 1);
            detail.SetColumnSpan(bindingGrid, 2);
            usageLabel = new Label
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(8),
                AutoEllipsis = true,
                Text = "Chọn một dòng để xem đơn giá dùng chung."
            };
            detail.Controls.Add(usageLabel, 6, 1);
            detail.SetColumnSpan(usageLabel, 2);

            var detailCommands = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };
            Button assignButton = Button("btnAssignEstimateNorm", "Gắn định mức", 108);
            Button clearButton = Button("btnClearEstimateNorm", "Dòng văn bản", 108);
            assignButton.Click += (sender, args) => Execute("Gắn định mức", AssignSelectedNorm);
            clearButton.Click += (sender, args) => ClearSelectedNorm();
            detailCommands.Controls.Add(assignButton);
            detailCommands.Controls.Add(clearButton);
            detail.Controls.Add(detailCommands, 0, 2);
            detail.SetColumnSpan(detailCommands, 8);
            split.Panel2.Controls.Add(detail);

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
            LoadSavedWorkspace();
        }

        public WorkbookEstimatePreview CurrentPreview { get; private set; }
        public WorkbookEstimateWorkspacePreview WorkspacePreview => workspacePreview;

        public bool SetCondition(string conditionCode)
        {
            try
            {
                WorkbookUnitRateContext context = WorkbookUnitRateService.LoadRequired(workbook);
                WorkbookEstimatePlan plan = WorkbookEstimateAppendixService.ImportLegacyPlan(workbook, context);
                string condition = (conditionCode ?? string.Empty).Trim();
                CurrentPreview = WorkbookEstimateAppendixService.Preview(
                    context,
                    plan,
                    condition.Length == 0 ? Array.Empty<string>() : new[] { condition });
                return CurrentPreview != null;
            }
            catch
            {
                CurrentPreview = null;
                return false;
            }
        }

        private void LoadSavedWorkspace()
        {
            EstimateWorkspace saved;
            if (!WorkbookEstimateWorkspaceService.TryLoad(workbook, out saved))
            {
                totalsLabel.Text = "0 dòng";
                statusLabel.Text = "Quét vùng Excel để bắt đầu; không cần lưu hồ sơ giá trước.";
                return;
            }
            currentWorkspace = saved;
            PopulateGrid(saved);
            sourceLabel.Text = saved.Source.WorksheetName + "!" + saved.Source.SourceAddress;
            RefreshPreview();
        }

        private void ScanRange()
        {
            Excel.Range range = null;
            Form host = FindForm();
            try
            {
                if (host != null)
                    host.Hide();
                object picked = workbook.Application.InputBox(
                    "Chọn vùng bảng Phụ lục dự toán, gồm cả dòng tiêu đề.",
                    "Quét Phụ lục dự toán",
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    8);
                range = picked as Excel.Range;
            }
            finally
            {
                if (host != null && !host.IsDisposed)
                {
                    host.Show();
                    host.Activate();
                }
            }
            if (range == null)
                return;
            try
            {
                EnsureRangeBelongsToWorkbook(range);
                using (var mapping = new FrmEstimateRangeMapping(range))
                {
                    if (mapping.ShowDialog(FindForm()) != DialogResult.OK)
                        return;
                    currentWorkspace = WorkbookEstimateRangeReader.Read(
                        workbook,
                        range,
                        mapping.BuildColumnMap(),
                        mapping.FirstRowIsHeader,
                        mapping.DefaultEnvironment,
                        mapping.DefaultAudience,
                        currentWorkspace);
                }
                PopulateGrid(currentWorkspace);
                sourceLabel.Text = currentWorkspace.Source.WorksheetName + "!" +
                    currentWorkspace.Source.SourceAddress;
                statusLabel.ForeColor = Color.DarkGoldenrod;
                statusLabel.Text = "Đã quét " + currentWorkspace.Rows.Count +
                    " dòng. Gắn định mức rồi bấm Lưu bố trí.";
                RefreshPreview();
            }
            finally
            {
                Release(range);
            }
        }

        private void EnsureRangeBelongsToWorkbook(Excel.Range range)
        {
            Excel.Worksheet sheet = null;
            Excel.Workbook parent = null;
            try
            {
                sheet = range.Worksheet;
                parent = sheet.Parent as Excel.Workbook;
                if (parent == null || !string.Equals(
                    parent.FullName,
                    workbook.FullName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("Vùng được chọn phải thuộc workbook dự toán đang mở.");
                }
            }
            finally
            {
                Release(parent);
                Release(sheet);
            }
        }

        private void SaveWorkspace(bool saveWorkbook)
        {
            EstimateWorkspace workspace = BuildWorkspaceFromGrid();
            WorkbookEstimateWorkspaceService.Save(workbook, workspace);
            if (saveWorkbook)
                workbook.Save();
            currentWorkspace = workspace;
            statusLabel.ForeColor = Color.DarkGreen;
            statusLabel.Text = "Đã lưu bố trí Phụ lục DT và " + workspace.Rows.Count + " dòng nguồn.";
        }

        private void GenerateSheets()
        {
            SaveWorkspace(false);
            RefreshPreview();
            if (workspacePreview == null || !workspacePreview.IsValid)
                throw new InvalidOperationException("Còn dòng chưa đủ định mức hoặc hồ sơ giá; hãy xử lý trước khi sinh sheet.");
            WorkbookGeneratedEstimateWriteResult result = WorkbookGeneratedEstimateWriter.Apply(
                workbook,
                workspacePreview);
            workbook.Save();
            statusLabel.ForeColor = Color.DarkGreen;
            statusLabel.Text = "Đã sinh " + result.RateCount + " bảng đơn giá trên " +
                result.WorksheetNames.Count + " sheet và link " + result.LinkedRowCount +
                " dòng Phụ lục DT.";
        }

        public void RefreshPreview()
        {
            if (lineGrid.Rows.Count == 0)
            {
                workspacePreview = null;
                totalsLabel.Text = "0 dòng";
                return;
            }
            currentWorkspace = BuildWorkspaceFromGrid();
            workspacePreview = WorkbookEstimateCalculationService.Preview(workbook, currentWorkspace);
            foreach (DataGridViewRow gridRow in lineGrid.Rows)
            {
                RowDraft draft = (RowDraft)gridRow.Tag;
                WorkbookEstimateRatePreview rate = workspacePreview.FindRateForRow(draft.RowId);
                if (draft.NormKey.Length == 0 && draft.VariantCode.Length == 0)
                {
                    SetStatus(gridRow, "Văn bản", string.Empty, 0, SystemColors.Window);
                    continue;
                }
                if (rate == null)
                {
                    string error = workspacePreview.Plan.Errors.FirstOrDefault(value =>
                        value.StartsWith(draft.RowId + ":", StringComparison.OrdinalIgnoreCase)) ??
                        "Không lập được khóa đơn giá.";
                    SetStatus(gridRow, "Thiếu dữ liệu", error, 0, Color.LemonChiffon);
                    continue;
                }
                if (!rate.IsValid)
                {
                    SetStatus(gridRow, "Lỗi", rate.Error, rate.Group.Rows.Count, Color.MistyRose);
                    gridRow.Cells["colWorkspaceRateId"].Value = rate.Group.Identity.RateId;
                    continue;
                }
                SetStatus(gridRow, "Hợp lệ", string.Empty, rate.Group.Rows.Count, Color.Honeydew);
                gridRow.Cells["colWorkspaceRateId"].Value = rate.Group.Identity.RateId;
            }
            int textRows = currentWorkspace.Rows.Count(row => row.IsTextRow);
            totalsLabel.Text = currentWorkspace.Rows.Count + " dòng | " +
                workspacePreview.Rates.Count + " đơn giá | " + textRows + " văn bản";
            if (workspacePreview.IsValid)
            {
                statusLabel.ForeColor = Color.DarkGreen;
                statusLabel.Text = "Tính thử đạt: " + workspacePreview.CalculatedRowCount +
                    " dòng dùng " + workspacePreview.Rates.Count + " đơn giá duy nhất.";
            }
            else
            {
                int rateErrors = workspacePreview.Rates.Count(rate => !rate.IsValid);
                statusLabel.ForeColor = Color.DarkGoldenrod;
                statusLabel.Text = workspacePreview.Plan.Errors.Count + " lỗi cấu hình, " +
                    rateErrors + " đơn giá chưa tính được.";
            }
            LoadSelectedDetails();
        }

        private void SetStatus(
            DataGridViewRow row,
            string status,
            string error,
            int usageCount,
            Color color)
        {
            row.Cells["colWorkspaceStatus"].Value = status;
            row.Cells["colWorkspaceUses"].Value = usageCount;
            row.Cells["colWorkspaceRateId"].Value = string.Empty;
            row.DefaultCellStyle.BackColor = color;
            foreach (DataGridViewCell cell in row.Cells)
                cell.ToolTipText = error;
        }

        private void PopulateGrid(EstimateWorkspace workspace)
        {
            loadingDetails = true;
            try
            {
                lineGrid.Rows.Clear();
                foreach (EstimateWorkspaceRow row in workspace.Rows)
                {
                    var draft = new RowDraft(row);
                    int index = lineGrid.Rows.Add(
                        string.Empty,
                        row.SourceRowHint,
                        row.WorkCode,
                        row.Description,
                        row.Unit,
                        row.QuantityFormulaLocal,
                        row.Quantity,
                        EnvironmentText(row.Environment),
                        AudienceText(row.LaborAudience),
                        row.NormKey,
                        row.VariantCode,
                        string.Empty,
                        0);
                    lineGrid.Rows[index].Tag = draft;
                }
                if (lineGrid.Rows.Count > 0)
                    lineGrid.CurrentCell = lineGrid.Rows[0].Cells["colWorkspaceDescription"];
            }
            finally
            {
                loadingDetails = false;
            }
            LoadSelectedDetails();
        }

        private EstimateWorkspace BuildWorkspaceFromGrid()
        {
            if (currentWorkspace == null)
                throw new InvalidOperationException("Chưa quét vùng Phụ lục DT.");
            return new EstimateWorkspace(
                currentWorkspace.Source,
                lineGrid.Rows.Cast<DataGridViewRow>()
                    .Where(row => !row.IsNewRow)
                    .Select(row => row.Tag as RowDraft)
                    .Where(draft => draft != null)
                    .Select(draft => draft.ToRow()),
                DateTime.UtcNow);
        }

        private void LoadSelectedDetails()
        {
            if (loadingDetails)
                return;
            DataGridViewRow selected = SelectedGridRow;
            loadingDetails = true;
            try
            {
                if (selected == null)
                {
                    usageLabel.Text = "Chọn một dòng để xem chi tiết.";
                    return;
                }
                RowDraft draft = selected.Tag as RowDraft;
                if (draft == null)
                {
                    usageLabel.Text = "Dòng đang được nạp; vui lòng chọn lại sau khi hoàn tất.";
                    return;
                }
                SelectValue(environmentInput, draft.Environment);
                SelectValue(audienceInput, draft.LaborAudience);
                SelectNorm(draft.NormKey);
                LoadVariantsAndConditions(draft);
                WorkbookEstimateRatePreview rate = workspacePreview?.FindRateForRow(draft.RowId);
                usageLabel.Text = rate == null
                    ? "Dòng văn bản hoặc chưa tính được đơn giá."
                    : rate.Group.Identity.RateId + "\n" +
                        rate.Group.Rows.Count + " công tác đang dùng chung\n" +
                        rate.Group.Identity.NormKey + " / " + rate.Group.Identity.VariantCode;
            }
            finally
            {
                loadingDetails = false;
            }
        }

        private void LoadVariantsAndConditions(RowDraft selectedDraft)
        {
            if (loadingDetails && selectedDraft == null)
                return;
            string selectedVariant = selectedDraft?.VariantCode ?? string.Empty;
            string[] selectedConditions = selectedDraft?.Conditions ?? Array.Empty<string>();
            UnitRateResourceBinding[] selectedBindings = selectedDraft?.Bindings ??
                Array.Empty<UnitRateResourceBinding>();
            variantInput.Items.Clear();
            conditionInput.Items.Clear();
            bindingGrid.Rows.Clear();
            NormOption option = normInput.SelectedItem as NormOption;
            if (option == null)
                return;
            foreach (string variant in option.Definition.Variants)
                variantInput.Items.Add(variant);
            if (variantInput.Items.Count > 0)
                variantInput.SelectedIndex = Math.Max(0, variantInput.Items.IndexOf(selectedVariant));
            foreach (string condition in option.Definition.Adjustments
                .Select(adjustment => adjustment.Condition)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal))
            {
                int index = conditionInput.Items.Add(condition);
                conditionInput.SetItemChecked(index, selectedConditions.Contains(condition, StringComparer.Ordinal));
            }
            foreach (string resourceCode in option.Definition.Rates
                .Select(rate => rate.ResourceCode)
                .Where(UnitRateCalculator.IsLogicalResource)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal))
            {
                UnitRateResourceBinding binding = selectedBindings.FirstOrDefault(value =>
                    string.Equals(value.ResourceCode, resourceCode, StringComparison.OrdinalIgnoreCase));
                bindingGrid.Rows.Add(
                    resourceCode,
                    binding?.PriceCode ?? string.Empty,
                    binding?.Reason ?? string.Empty);
            }
        }

        private void AssignSelectedNorm()
        {
            DataGridViewRow selected = SelectedGridRow;
            if (selected == null)
                return;
            NormOption norm = normInput.SelectedItem as NormOption;
            string variant = Convert.ToString(variantInput.SelectedItem, CultureInfo.CurrentCulture) ?? string.Empty;
            if (norm == null || variant.Length == 0)
                throw new InvalidOperationException("Phải chọn định mức và mã chi tiết.");
            RowDraft draft = (RowDraft)selected.Tag;
            draft.NormKey = norm.Definition.Key;
            draft.VariantCode = variant;
            draft.Environment = SelectedValue<EstimateWorkEnvironment>(environmentInput);
            draft.LaborAudience = SelectedValue<MachineRateAudience>(audienceInput);
            draft.Conditions = conditionInput.CheckedItems.Cast<object>()
                .Select(value => Convert.ToString(value, CultureInfo.CurrentCulture))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToArray();
            draft.Bindings = bindingGrid.Rows.Cast<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .Select(row => new UnitRateResourceBinding(
                    CellText(row, "colBindingResource"),
                    CellText(row, "colBindingPrice"),
                    CellText(row, "colBindingReason")))
                .ToArray();
            UpdateAssignmentCells(selected, draft);
            RefreshPreview();
        }

        private void ClearSelectedNorm()
        {
            DataGridViewRow selected = SelectedGridRow;
            if (selected == null)
                return;
            RowDraft draft = (RowDraft)selected.Tag;
            draft.NormKey = string.Empty;
            draft.VariantCode = string.Empty;
            draft.Conditions = Array.Empty<string>();
            draft.Bindings = Array.Empty<UnitRateResourceBinding>();
            UpdateAssignmentCells(selected, draft);
            RefreshPreview();
        }

        private static void UpdateAssignmentCells(DataGridViewRow row, RowDraft draft)
        {
            row.Cells["colWorkspaceEnvironment"].Value = EnvironmentText(draft.Environment);
            row.Cells["colWorkspaceAudience"].Value = AudienceText(draft.LaborAudience);
            row.Cells["colWorkspaceNorm"].Value = draft.NormKey;
            row.Cells["colWorkspaceVariant"].Value = draft.VariantCode;
        }

        private void SelectNorm(string normKey)
        {
            normInput.SelectedIndex = -1;
            for (int index = 0; index < normInput.Items.Count; index++)
            {
                NormOption option = (NormOption)normInput.Items[index];
                if (string.Equals(option.Definition.Key, normKey, StringComparison.Ordinal))
                {
                    normInput.SelectedIndex = index;
                    return;
                }
            }
        }

        private DataGridViewRow SelectedGridRow =>
            lineGrid.CurrentRow != null && !lineGrid.CurrentRow.IsNewRow
                ? lineGrid.CurrentRow
                : null;

        private DataGridView CreateGrid()
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                DefaultCellStyle = { FormatProvider = numberCulture }
            };
            grid.Columns.Add(TextColumn("colWorkspaceStatus", "Trạng thái", 82));
            grid.Columns.Add(NumberColumn("colWorkspaceRow", "Hàng", 50, "0"));
            grid.Columns.Add(TextColumn("colWorkspaceCode", "Số hiệu", 82));
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colWorkspaceDescription",
                HeaderText = "Nội dung công tác",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 190
            });
            grid.Columns.Add(TextColumn("colWorkspaceUnit", "ĐVT", 58));
            grid.Columns.Add(TextColumn("colWorkspaceFormula", "Công thức KL", 105));
            grid.Columns.Add(NumberColumn("colWorkspaceQuantity", "Khối lượng", 78, "#,##0.##"));
            grid.Columns.Add(TextColumn("colWorkspaceEnvironment", "Cạn/Nước", 78));
            grid.Columns.Add(TextColumn("colWorkspaceAudience", "Đối tượng lương", 112));
            grid.Columns.Add(TextColumn("colWorkspaceNorm", "Định mức", 105));
            grid.Columns.Add(TextColumn("colWorkspaceVariant", "Mã chi tiết", 94));
            grid.Columns.Add(TextColumn("colWorkspaceRateId", "Mã đơn giá", 112));
            grid.Columns.Add(NumberColumn("colWorkspaceUses", "Dùng chung", 64, "0"));
            return grid;
        }

        private void Execute(string operation, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex,
                    operation,
                    "DT-701",
                    operation,
                    workbook,
                    "EstimateAppendix");
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
                MessageBox.Show(ex.Message, operation, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static Button Button(string name, string text, int width)
        {
            return new Button { Name = name, Text = text, Width = width, Height = 30 };
        }

        private static ComboBox DropDown(string name)
        {
            return new ComboBox
            {
                Name = name,
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                IntegralHeight = false,
                MaxDropDownItems = 16
            };
        }

        private static Label DetailLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
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
                DefaultCellStyle =
                {
                    Alignment = DataGridViewContentAlignment.MiddleRight,
                    Format = format
                }
            };
        }

        private static void SelectValue<T>(ComboBox combo, T value)
        {
            for (int index = 0; index < combo.Items.Count; index++)
            {
                Option<T> option = combo.Items[index] as Option<T>;
                if (option != null && EqualityComparer<T>.Default.Equals(option.Value, value))
                {
                    combo.SelectedIndex = index;
                    return;
                }
            }
            combo.SelectedIndex = combo.Items.Count > 0 ? 0 : -1;
        }

        private static T SelectedValue<T>(ComboBox combo)
        {
            Option<T> option = combo.SelectedItem as Option<T>;
            if (option == null)
                throw new InvalidOperationException("Chưa chọn " + combo.Name + ".");
            return option.Value;
        }

        private static string EnvironmentText(EstimateWorkEnvironment environment)
        {
            return environment == EstimateWorkEnvironment.Land ? "Cạn" : "Nước";
        }

        private static string AudienceText(MachineRateAudience audience)
        {
            return audience == MachineRateAudience.StateBudgetSalary ? "HLNS" : "KHLNS";
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private static string CellText(DataGridViewRow row, string column)
        {
            return (Convert.ToString(row.Cells[column].Value, CultureInfo.CurrentCulture) ?? string.Empty).Trim();
        }

        private sealed class NormOption
        {
            internal NormOption(NormDefinition definition)
            {
                Definition = definition;
            }

            internal NormDefinition Definition { get; }
            public override string ToString() => Definition.Key + " - " + Definition.Title;
        }

        private sealed class Option<T>
        {
            internal Option(string text, T value)
            {
                Text = text;
                Value = value;
            }

            internal string Text { get; }
            internal T Value { get; }
            public override string ToString() => Text;
        }

        private sealed class RowDraft
        {
            internal RowDraft(EstimateWorkspaceRow row)
            {
                RowId = row.RowId;
                BindingName = row.BindingName;
                SourceRow = row.SourceRowHint;
                WorkCode = row.WorkCode;
                Description = row.Description;
                Unit = row.Unit;
                QuantityFormula = row.QuantityFormulaLocal;
                Quantity = row.Quantity;
                AcceptedFormula = row.AcceptedQuantityFormulaLocal;
                AcceptedQuantity = row.AcceptedQuantity;
                Environment = row.Environment;
                LaborAudience = row.LaborAudience;
                NormKey = row.NormKey;
                VariantCode = row.VariantCode;
                Conditions = row.Conditions.ToArray();
                Bindings = row.Bindings.ToArray();
            }

            internal string RowId { get; }
            internal string BindingName { get; }
            internal int SourceRow { get; }
            internal string WorkCode { get; }
            internal string Description { get; }
            internal string Unit { get; }
            internal string QuantityFormula { get; }
            internal decimal Quantity { get; }
            internal string AcceptedFormula { get; }
            internal decimal AcceptedQuantity { get; }
            internal EstimateWorkEnvironment Environment { get; set; }
            internal MachineRateAudience LaborAudience { get; set; }
            internal string NormKey { get; set; }
            internal string VariantCode { get; set; }
            internal string[] Conditions { get; set; }
            internal UnitRateResourceBinding[] Bindings { get; set; }

            internal EstimateWorkspaceRow ToRow()
            {
                return new EstimateWorkspaceRow(
                    RowId,
                    BindingName,
                    SourceRow,
                    WorkCode,
                    Description,
                    Unit,
                    QuantityFormula,
                    Quantity,
                    AcceptedFormula,
                    AcceptedQuantity,
                    Environment,
                    LaborAudience,
                    NormKey,
                    VariantCode,
                    Conditions,
                    Bindings);
            }
        }
    }
}
