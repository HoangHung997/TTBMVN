using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class FrmWorksheetRoleMapping : Form, IWorkbookSheetListObserver
    {
        private readonly Excel.Workbook workbook;
        private readonly WorkbookSheetChangeCoordinator coordinator;
        private readonly DataGridView grid;
        private readonly DataGridViewComboBoxColumn sheetColumn;
        private readonly Label statusLabel;
        private IDisposable subscription;
        private IReadOnlyList<WorkbookSheetDescriptor> availableSheets =
            new WorkbookSheetDescriptor[0];

        public FrmWorksheetRoleMapping(
            Excel.Workbook workbook,
            WorkbookSheetChangeCoordinator coordinator)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.coordinator = coordinator;

            Text = "Thiết lập vai trò sheet";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(700, 430);
            Size = new Size(760, 480);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;

            var header = new Label
            {
                Dock = DockStyle.Top,
                Height = 46,
                Padding = new Padding(12, 12, 12, 6),
                Text = "Chọn đúng một sheet cho mỗi vai trò nghiệp vụ.",
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold)
            };

            grid = new DataGridView
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
            grid.DataError += Grid_DataError;
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Role",
                HeaderText = "Vai trò",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 52
            });
            sheetColumn = new DataGridViewComboBoxColumn
            {
                Name = "Sheet",
                HeaderText = "Sheet",
                DisplayMember = nameof(WorkbookSheetDescriptor.Name),
                ValueMember = nameof(WorkbookSheetDescriptor.Key),
                ValueType = typeof(string),
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 48,
                FlatStyle = FlatStyle.Flat
            };
            grid.Columns.Add(sheetColumn);

            statusLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 42,
                Padding = new Padding(12, 8, 12, 4),
                ForeColor = Color.Firebrick,
                AutoEllipsis = true
            };

            var saveButton = new Button
            {
                Text = "Lưu ánh xạ",
                AutoSize = true,
                Padding = new Padding(10, 2, 10, 2)
            };
            saveButton.Click += SaveButton_Click;
            var cancelButton = new Button
            {
                Text = "Đóng",
                DialogResult = DialogResult.Cancel,
                AutoSize = true,
                Padding = new Padding(10, 2, 10, 2)
            };
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 52,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(8),
                WrapContents = false
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(saveButton);

            Controls.Add(grid);
            Controls.Add(statusLabel);
            Controls.Add(header);
            Controls.Add(buttons);
            AcceptButton = saveButton;
            CancelButton = cancelButton;

            BuildRoleRows();
            OnWorkbookSheetsChanged(WorkbookSheetChangeCoordinator.CaptureSnapshot(workbook));
            ApplyExistingRolesAndSuggestions();
            FormClosed += FrmWorksheetRoleMapping_FormClosed;
            if (coordinator != null)
                subscription = coordinator.Subscribe(workbook, this);
        }

        public void OnWorkbookSheetsChanged(IReadOnlyList<WorkbookSheetDescriptor> sheets)
        {
            if (IsDisposed || Disposing)
                return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<IReadOnlyList<WorkbookSheetDescriptor>>(OnWorkbookSheetsChanged), sheets);
                return;
            }

            var previousKeys = new Dictionary<WorksheetRole, string>();
            foreach (DataGridViewRow row in grid.Rows)
                previousKeys[(WorksheetRole)row.Tag] = Convert.ToString(row.Cells[1].Value) ?? string.Empty;

            availableSheets = WorkbookSheetList.Copy(sheets);
            sheetColumn.Items.Clear();
            foreach (WorkbookSheetDescriptor sheet in availableSheets)
                sheetColumn.Items.Add(sheet);

            foreach (DataGridViewRow row in grid.Rows)
            {
                WorksheetRole role = (WorksheetRole)row.Tag;
                string key = previousKeys[role];
                row.Cells[1].Value = availableSheets.Any(sheet =>
                    string.Equals(sheet.Key, key, StringComparison.OrdinalIgnoreCase))
                    ? key
                    : null;
            }
        }

        private void BuildRoleRows()
        {
            foreach (WorksheetRole role in WorksheetRoleCatalog.All)
            {
                int index = grid.Rows.Add(GetRoleLabel(role), null);
                grid.Rows[index].Tag = role;
            }
        }

        private void ApplyExistingRolesAndSuggestions()
        {
            var selectedByRole = new Dictionary<WorksheetRole, string>();
            foreach (WorksheetRoleAssignment assignment in WorksheetRoleService.ReadAssignments(workbook))
            {
                WorksheetRole role;
                if (!WorksheetRoleCatalog.TryParse(assignment.RoleId, out role))
                    continue;
                WorkbookSheetDescriptor sheet = availableSheets.FirstOrDefault(item =>
                    string.Equals(item.Name, assignment.SheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet != null)
                    selectedByRole[role] = sheet.Key;
            }

            foreach (WorksheetRoleMappingEntry suggestion in
                WorksheetRoleMappingSuggester.Suggest(availableSheets))
            {
                if (!selectedByRole.ContainsKey(suggestion.Role) && suggestion.SheetKey.Length > 0)
                    selectedByRole[suggestion.Role] = suggestion.SheetKey;
            }

            foreach (DataGridViewRow row in grid.Rows)
            {
                WorksheetRole role = (WorksheetRole)row.Tag;
                string key;
                if (selectedByRole.TryGetValue(role, out key))
                    row.Cells[1].Value = key;
            }
        }

        private IReadOnlyList<WorksheetRoleMappingEntry> ReadMapping()
        {
            var result = new List<WorksheetRoleMappingEntry>();
            foreach (DataGridViewRow row in grid.Rows)
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
                IReadOnlyList<WorksheetRoleMappingEntry> mapping = ReadMapping();
                WorksheetRoleMappingValidationResult validation =
                    WorksheetRoleMappingValidator.Validate(mapping, availableSheets);
                if (!validation.IsValid)
                {
                    statusLabel.Text = string.Join(" ", validation.Issues.Select(issue => issue.Message));
                    return;
                }

                WorksheetRoleService.ApplyMapping(workbook, mapping);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Apply worksheet role mapping");
                statusLabel.Text = ex.Message;
            }
        }

        private void Grid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
        }

        private void FrmWorksheetRoleMapping_FormClosed(object sender, FormClosedEventArgs e)
        {
            subscription?.Dispose();
            subscription = null;
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

