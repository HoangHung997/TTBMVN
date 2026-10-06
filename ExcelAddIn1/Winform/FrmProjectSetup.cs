using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class FrmProjectSetup : Form, IWorkbookSheetListObserver
    {
        private readonly Excel.Workbook workbook;
        private readonly WorkbookSheetChangeCoordinator coordinator;
        private readonly IReadOnlyList<RegulationPackage> packages;
        private readonly TextBox projectIdTextBox;
        private readonly DateTimePicker preparedDatePicker;
        private readonly CheckBox approvalCheckBox;
        private readonly DateTimePicker approvalDatePicker;
        private readonly DateTimePicker evaluationDatePicker;
        private readonly DateTimePicker priceDatePicker;
        private readonly TextBox priceProfileTextBox;
        private readonly Label packageValueLabel;
        private readonly Label packageReasonLabel;
        private readonly DataGridView mappingGrid;
        private readonly DataGridViewComboBoxColumn sheetColumn;
        private readonly Label statusLabel;
        private readonly Button saveButton;
        private IDisposable subscription;
        private IReadOnlyList<WorkbookSheetDescriptor> availableSheets =
            new WorkbookSheetDescriptor[0];
        private bool initializing;

        public FrmProjectSetup(
            Excel.Workbook workbook,
            WorkbookSheetChangeCoordinator coordinator)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.coordinator = coordinator;
            packages = RegulationPackageBootstrapService.LoadPreferredPackages();

            Text = "Thiết lập dự án dự toán";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(850, 650);
            Size = new Size(920, 720);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;

            var header = new Label
            {
                Dock = DockStyle.Top,
                Height = 48,
                Padding = new Padding(14, 13, 14, 6),
                Text = "Hồ sơ, căn cứ pháp lý, mốc giá và vai trò sheet",
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold)
            };

            var profileTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 178,
                Padding = new Padding(12, 4, 12, 4),
                ColumnCount = 4,
                RowCount = 4
            };
            profileTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 125));
            profileTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            profileTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 125));
            profileTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            for (int row = 0; row < 4; row++)
                profileTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            projectIdTextBox = CreateTextBox();
            preparedDatePicker = CreateDatePicker();
            approvalCheckBox = new CheckBox
            {
                Text = "Đã phê duyệt",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            approvalDatePicker = CreateDatePicker();
            evaluationDatePicker = CreateDatePicker();
            priceDatePicker = CreateDatePicker();
            priceProfileTextBox = CreateTextBox();

            AddProfileField(profileTable, "Mã dự án", projectIdTextBox, 0, 0);
            AddProfileField(profileTable, "Ngày lập", preparedDatePicker, 2, 0);
            profileTable.Controls.Add(approvalCheckBox, 0, 1);
            profileTable.Controls.Add(approvalDatePicker, 1, 1);
            AddProfileField(profileTable, "Ngày đánh giá", evaluationDatePicker, 2, 1);
            AddProfileField(profileTable, "Mốc giá", priceDatePicker, 0, 2);
            AddProfileField(profileTable, "Hồ sơ giá", priceProfileTextBox, 2, 2);

            packageValueLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                AutoEllipsis = true
            };
            packageReasonLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.GrayText,
                AutoEllipsis = true
            };
            profileTable.Controls.Add(CreateFieldLabel("Căn cứ"), 0, 3);
            profileTable.Controls.Add(packageValueLabel, 1, 3);
            profileTable.Controls.Add(packageReasonLabel, 2, 3);
            profileTable.SetColumnSpan(packageReasonLabel, 2);

            mappingGrid = new DataGridView
            {
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
            mappingGrid.DataError += (sender, args) => args.ThrowException = false;
            mappingGrid.CurrentCellDirtyStateChanged += (sender, args) =>
            {
                if (mappingGrid.IsCurrentCellDirty)
                    mappingGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            mappingGrid.CellValueChanged += (sender, args) => RefreshPlanPreview();
            mappingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Role",
                HeaderText = "Vai trò nghiệp vụ",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 52
            });
            sheetColumn = new DataGridViewComboBoxColumn
            {
                Name = "Sheet",
                HeaderText = "Sheet trong workbook",
                DisplayMember = nameof(WorkbookSheetDescriptor.Name),
                ValueMember = nameof(WorkbookSheetDescriptor.Key),
                ValueType = typeof(string),
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 48,
                FlatStyle = FlatStyle.Flat
            };
            mappingGrid.Columns.Add(sheetColumn);

            var mappingHeader = new Label
            {
                Dock = DockStyle.Top,
                Height = 34,
                Padding = new Padding(12, 8, 12, 2),
                Text = "Ánh xạ sheet theo vai trò",
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold)
            };

            statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 58,
                Padding = new Padding(12, 8, 12, 4),
                ForeColor = Color.Firebrick,
                AutoEllipsis = true
            };
            saveButton = new Button
            {
                Text = "Lưu và mở dự toán",
                AutoSize = true,
                Padding = new Padding(10, 3, 10, 3)
            };
            saveButton.Click += SaveButton_Click;
            var cancelButton = new Button
            {
                Text = "Hủy",
                DialogResult = DialogResult.Cancel,
                AutoSize = true,
                Padding = new Padding(10, 3, 10, 3)
            };
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 54,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(8),
                WrapContents = false
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(saveButton);

            Controls.Add(mappingGrid);
            Controls.Add(mappingHeader);
            Controls.Add(profileTable);
            Controls.Add(header);
            Controls.Add(statusLabel);
            Controls.Add(buttons);
            AcceptButton = saveButton;
            CancelButton = cancelButton;

            BuildRoleRows();
            initializing = true;
            OnWorkbookSheetsChanged(WorkbookSheetChangeCoordinator.CaptureSnapshot(workbook));
            LoadProfile();
            ApplyExistingRolesAndSuggestions();
            initializing = false;
            WireProfileEvents();
            RefreshPlanPreview();

            if (coordinator != null)
                subscription = coordinator.Subscribe(workbook, this);
            FormClosed += (sender, args) =>
            {
                subscription?.Dispose();
                subscription = null;
            };
        }

        public void OnWorkbookSheetsChanged(IReadOnlyList<WorkbookSheetDescriptor> sheets)
        {
            if (IsDisposed || Disposing)
                return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<IReadOnlyList<WorkbookSheetDescriptor>>(
                    OnWorkbookSheetsChanged), sheets);
                return;
            }

            var previousKeys = new Dictionary<WorksheetRole, string>();
            foreach (DataGridViewRow row in mappingGrid.Rows)
            {
                previousKeys[(WorksheetRole)row.Tag] =
                    Convert.ToString(row.Cells[1].Value) ?? string.Empty;
            }

            availableSheets = WorkbookSheetList.Copy(sheets);
            sheetColumn.Items.Clear();
            foreach (WorkbookSheetDescriptor sheet in availableSheets)
                sheetColumn.Items.Add(sheet);

            foreach (DataGridViewRow row in mappingGrid.Rows)
            {
                WorksheetRole role = (WorksheetRole)row.Tag;
                string key = previousKeys[role];
                row.Cells[1].Value = availableSheets.Any(sheet => string.Equals(
                    sheet.Key,
                    key,
                    StringComparison.OrdinalIgnoreCase)) ? key : null;
            }
            RefreshPlanPreview();
        }

        private void LoadProfile()
        {
            ProjectProfile profile;
            DateTime today = DateTime.Today;
            if (WorkbookProjectProfileService.TryLoad(workbook, out profile))
            {
                projectIdTextBox.Text = profile.ProjectId;
                preparedDatePicker.Value = profile.PreparedDate.Value.Date;
                approvalCheckBox.Checked = profile.ApprovalDate.HasValue;
                if (profile.ApprovalDate.HasValue)
                    approvalDatePicker.Value = profile.ApprovalDate.Value.Date;
                evaluationDatePicker.Value = MaxDate(
                    today,
                    profile.PreparedDate.Value,
                    profile.ApprovalDate);
                priceDatePicker.Value = profile.PriceDate.Value.Date;
                priceProfileTextBox.Text = profile.PriceProfileId;
            }
            else
            {
                projectIdTextBox.Text = Path.GetFileNameWithoutExtension(workbook.Name);
                preparedDatePicker.Value = today;
                approvalCheckBox.Checked = false;
                approvalDatePicker.Value = today;
                evaluationDatePicker.Value = today;
                priceDatePicker.Value = today;
                priceProfileTextBox.Text = "PRICE-" + today.ToString("yyyy-MM");
            }
            approvalDatePicker.Enabled = approvalCheckBox.Checked;
        }

        private void ApplyExistingRolesAndSuggestions()
        {
            IReadOnlyList<WorksheetRoleMappingEntry> mapping =
                WorksheetRoleMappingBuilder.FromAssignments(
                    WorksheetRoleService.ReadAssignments(workbook),
                    availableSheets);
            foreach (DataGridViewRow row in mappingGrid.Rows)
            {
                WorksheetRoleMappingEntry entry = mapping.First(item =>
                    item.Role == (WorksheetRole)row.Tag);
                if (entry.SheetKey.Length > 0)
                    row.Cells[1].Value = entry.SheetKey;
            }
        }

        private void WireProfileEvents()
        {
            projectIdTextBox.TextChanged += (sender, args) => RefreshPlanPreview();
            preparedDatePicker.ValueChanged += (sender, args) => RefreshPlanPreview();
            approvalCheckBox.CheckedChanged += (sender, args) =>
            {
                approvalDatePicker.Enabled = approvalCheckBox.Checked;
                RefreshPlanPreview();
            };
            approvalDatePicker.ValueChanged += (sender, args) => RefreshPlanPreview();
            evaluationDatePicker.ValueChanged += (sender, args) => RefreshPlanPreview();
            priceDatePicker.ValueChanged += (sender, args) => RefreshPlanPreview();
            priceProfileTextBox.TextChanged += (sender, args) => RefreshPlanPreview();
        }

        private void RefreshPlanPreview()
        {
            if (initializing || saveButton == null)
                return;

            try
            {
                ProjectSetupPlanningResult result = CreatePlan();
                if (result.PackageResolution != null &&
                    result.PackageResolution.IsSuccess)
                {
                    RegulationPackage package = result.PackageResolution.Package;
                    packageValueLabel.Text = package.PackageId + " v" + package.DataVersion;
                    packageReasonLabel.Text = result.PackageResolution.Reason;
                }
                else
                {
                    packageValueLabel.Text = "Chưa xác định";
                    packageReasonLabel.Text = result.PackageResolution?.Reason ?? string.Empty;
                }

                saveButton.Enabled = result.IsValid;
                statusLabel.ForeColor = result.IsValid ? Color.DarkGreen : Color.Firebrick;
                statusLabel.Text = result.IsValid
                    ? "Thiết lập hợp lệ. Workbook chỉ được thay đổi khi bấm Lưu và mở dự toán."
                    : string.Join(" ", result.Errors);
            }
            catch (Exception ex)
            {
                saveButton.Enabled = false;
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
            }
        }

        private ProjectSetupPlanningResult CreatePlan()
        {
            var draft = new ProjectSetupDraft(
                projectIdTextBox.Text,
                preparedDatePicker.Value.Date,
                approvalCheckBox.Checked ? approvalDatePicker.Value.Date : (DateTime?)null,
                evaluationDatePicker.Value.Date,
                priceDatePicker.Value.Date,
                priceProfileTextBox.Text,
                string.Empty,
                ReadMapping());
            return ProjectSetupPlanner.CreatePlan(
                draft,
                availableSheets,
                packages,
                RegulationTransitionRuleCatalog.All);
        }

        private IReadOnlyList<WorksheetRoleMappingEntry> ReadMapping()
        {
            var result = new List<WorksheetRoleMappingEntry>();
            foreach (DataGridViewRow row in mappingGrid.Rows)
            {
                WorksheetRole role = (WorksheetRole)row.Tag;
                string key = Convert.ToString(row.Cells[1].Value) ?? string.Empty;
                WorkbookSheetDescriptor sheet = availableSheets.FirstOrDefault(item =>
                    string.Equals(item.Key, key, StringComparison.OrdinalIgnoreCase));
                result.Add(new WorksheetRoleMappingEntry(
                    role,
                    sheet?.Key ?? key,
                    sheet?.Name ?? string.Empty));
            }
            return result.AsReadOnly();
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            try
            {
                ProjectSetupPlanningResult result = CreatePlan();
                if (!result.IsValid)
                {
                    RefreshPlanPreview();
                    return;
                }

                WorkbookProjectSetupService.Commit(workbook, result.Plan);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Commit project setup", "DT-401", "commit", workbook, string.Empty);
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
            }
        }

        private void BuildRoleRows()
        {
            foreach (WorksheetRole role in WorksheetRoleCatalog.All)
            {
                int index = mappingGrid.Rows.Add(GetRoleLabel(role), null);
                mappingGrid.Rows[index].Tag = role;
            }
        }

        private static TextBox CreateTextBox()
        {
            return new TextBox { Dock = DockStyle.Fill, Margin = new Padding(3, 7, 8, 6) };
        }

        private static DateTimePicker CreateDatePicker()
        {
            return new DateTimePicker
            {
                Dock = DockStyle.Fill,
                Format = DateTimePickerFormat.Short,
                Margin = new Padding(3, 6, 8, 6)
            };
        }

        private static Label CreateFieldLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static void AddProfileField(
            TableLayoutPanel table,
            string label,
            Control control,
            int column,
            int row)
        {
            table.Controls.Add(CreateFieldLabel(label), column, row);
            table.Controls.Add(control, column + 1, row);
        }

        private static DateTime MaxDate(
            DateTime first,
            DateTime second,
            DateTime? third)
        {
            DateTime result = first.Date > second.Date ? first.Date : second.Date;
            if (third.HasValue && third.Value.Date > result)
                result = third.Value.Date;
            return result;
        }

        private static string GetRoleLabel(WorksheetRole role)
        {
            switch (role)
            {
                case WorksheetRole.ResourcePrices: return "Giá vật liệu - nhân công - máy";
                case WorksheetRole.UnitRateLand: return "Đơn giá trên cạn";
                case WorksheetRole.UnitRateWater: return "Đơn giá dưới nước";
                case WorksheetRole.EstimateAppendix: return "Phụ lục dự toán";
                case WorksheetRole.CostSummary: return "Tổng hợp kinh phí";
                case WorksheetRole.NormLookupView: return "Tra cứu định mức";
                case WorksheetRole.CostRuleView: return "Quy tắc chi phí";
                default: return WorksheetRoleCatalog.ToId(role);
            }
        }
    }
}
