using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class WorkbookValidationControl : UserControl
    {
        private readonly Excel.Workbook workbook;
        private readonly DataGridView issueGrid;
        private readonly Label statusLabel;

        public WorkbookValidationControl(Excel.Workbook workbook)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            Padding = new Padding(14);

            var commands = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 42,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            var scanButton = new Button
            {
                Name = "btnScanWorkbook",
                Text = "Kiểm tra workbook",
                Width = 136,
                Height = 30
            };
            var navigateButton = new Button
            {
                Name = "btnNavigateValidationIssue",
                Text = "Đi đến ô",
                Width = 86,
                Height = 30
            };
            scanButton.Click += (sender, args) => RunScan();
            navigateButton.Click += (sender, args) => NavigateSelected();
            commands.Controls.Add(scanButton);
            commands.Controls.Add(navigateButton);
            commands.Controls.Add(new Label
            {
                Text = "Validation và đối soát",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(12, 7, 0, 0)
            });

            issueGrid = new DataGridView
            {
                Name = "dgvWorkbookValidation",
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                MultiSelect = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            issueGrid.Columns.Add(Column("severity", "Mức", 70));
            issueGrid.Columns.Add(Column("code", "Mã lỗi", 116));
            issueGrid.Columns.Add(Column("location", "Vị trí", 142));
            issueGrid.Columns.Add(Column("subject", "Đối tượng", 132));
            DataGridViewTextBoxColumn message = Column("message", "Nội dung", 250);
            message.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            issueGrid.Columns.Add(message);
            issueGrid.Columns.Add(Column("remediation", "Cách xử lý", 210));
            issueGrid.CellDoubleClick += (sender, args) => NavigateSelected();

            statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 34,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(35, 92, 55)
            };

            Controls.Add(issueGrid);
            Controls.Add(statusLabel);
            Controls.Add(commands);
            statusLabel.Text = "Bấm Kiểm tra workbook để quét dữ liệu hiện tại.";
        }

        public WorkbookValidationReport CurrentReport { get; private set; }

        public WorkbookValidationReport RunScan()
        {
            try
            {
                CurrentReport = WorkbookValidationService.Scan(workbook);
                issueGrid.Rows.Clear();
                foreach (WorkbookValidationIssue issue in CurrentReport.Issues)
                {
                    int index = issueGrid.Rows.Add(
                        SeverityText(issue),
                        issue.Code,
                        IssueLocation(issue),
                        issue.Subject,
                        issue.Message,
                        issue.Remediation);
                    DataGridViewRow row = issueGrid.Rows[index];
                    row.Tag = issue;
                    if (issue.IsBlocking)
                        row.DefaultCellStyle.ForeColor = Color.FromArgb(165, 42, 42);
                    else if (issue.Severity == WorkbookValidationSeverity.Warning || issue.IsInherited)
                        row.DefaultCellStyle.ForeColor = Color.FromArgb(145, 91, 20);
                }
                statusLabel.Text = CurrentReport.IsValid
                    ? "Đạt: 0 lỗi chặn, " + CurrentReport.WarningCount.ToString(CultureInfo.CurrentCulture) + " cảnh báo."
                    : "Chưa đạt: " + CurrentReport.ErrorCount.ToString(CultureInfo.CurrentCulture) +
                        " lỗi chặn, " + CurrentReport.WarningCount.ToString(CultureInfo.CurrentCulture) + " cảnh báo.";
                return CurrentReport;
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Scan workbook validation", "DT-503", "scan", workbook, string.Empty);
                statusLabel.Text = ex.Message;
                CurrentReport = null;
                return null;
            }
        }

        public bool NavigateSelected()
        {
            if (issueGrid.SelectedRows.Count == 0)
                return false;
            WorkbookValidationIssue issue = issueGrid.SelectedRows[0].Tag as WorkbookValidationIssue;
            try
            {
                return WorkbookValidationService.NavigateTo(workbook, issue);
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex,
                    "Navigate workbook validation issue",
                    "DT-503",
                    "navigate",
                    workbook,
                    issue?.WorksheetRoleId);
                statusLabel.Text = ex.Message;
                return false;
            }
        }

        private static DataGridViewTextBoxColumn Column(string name, string text, int width)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = text,
                Width = width,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private static string IssueLocation(WorkbookValidationIssue issue)
        {
            string owner = issue.WorksheetRoleId.Length > 0
                ? issue.WorksheetRoleId
                : issue.WorksheetCodeName;
            return issue.Address.Length == 0 ? owner : owner + "!" + issue.Address;
        }

        private static string SeverityText(WorkbookValidationIssue issue)
        {
            if (issue.IsInherited)
                return "Kế thừa";
            switch (issue.Severity)
            {
                case WorkbookValidationSeverity.Error: return "Lỗi";
                case WorkbookValidationSeverity.Warning: return "Cảnh báo";
                default: return "Thông tin";
            }
        }
    }
}
