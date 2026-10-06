using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal sealed class EstimateCostSummaryPaneView : UserControl
    {
        private static readonly Color Green = Color.FromArgb(0, 137, 70);
        private static readonly Color GreenDark = Color.FromArgb(0, 103, 55);
        private static readonly Color GreenSoft = Color.FromArgb(235, 248, 239);
        private static readonly Color Blue = Color.FromArgb(45, 112, 229);
        private static readonly Color BlueSoft = Color.FromArgb(237, 244, 255);
        private static readonly Color Amber = Color.FromArgb(244, 166, 35);
        private static readonly Color AmberSoft = Color.FromArgb(255, 248, 226);
        private static readonly Color Red = Color.FromArgb(220, 53, 69);
        private static readonly Color Border = Color.FromArgb(222, 228, 223);
        private static readonly Color TextDark = Color.FromArgb(33, 43, 54);
        private static readonly Color TextMuted = Color.FromArgb(92, 103, 112);

        private readonly Excel.Workbook workbook;
        private readonly Action backAction;
        private readonly Label totalValue;
        private readonly Label ratedValue;
        private readonly Label warningValue;
        private readonly Label formulaErrorValue;
        private readonly FlowLayoutPanel issuePanel;
        private readonly Label statusLabel;
        private WorkbookEstimateV2CostPreview preview;

        internal EstimateCostSummaryPaneView(
            Excel.Workbook workbook,
            Action backAction)
        {
            this.workbook = workbook ??
                throw new ArgumentNullException(nameof(workbook));
            this.backAction = backAction;

            Dock = DockStyle.Fill;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;

            var root = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White
            };
            Controls.Add(root);

            var content = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                BackColor = Color.White,
                Padding = new Padding(14, 10, 14, 14),
                ColumnCount = 1
            };
            root.Controls.Add(content);

            content.Controls.Add(
                BuildHeader(),
                0,
                content.RowCount++);

            var metrics = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 98,
                ColumnCount = 4,
                Margin = new Padding(0, 4, 0, 10)
            };
            for (int index = 0; index < 4; index++)
                metrics.ColumnStyles.Add(
                    new ColumnStyle(SizeType.Percent, 25));

            totalValue = AddMetric(
                metrics,
                0,
                "Tổng công tác",
                EstimateUiIconKind.Clipboard,
                Green);
            ratedValue = AddMetric(
                metrics,
                1,
                "Đủ đơn giá",
                EstimateUiIconKind.Check,
                Green);
            warningValue = AddMetric(
                metrics,
                2,
                "Cảnh báo",
                EstimateUiIconKind.Warning,
                Amber);
            formulaErrorValue = AddMetric(
                metrics,
                3,
                "Lỗi công thức",
                EstimateUiIconKind.Error,
                Red);

            content.Controls.Add(
                metrics,
                0,
                content.RowCount++);

            var update = new Panel
            {
                Dock = DockStyle.Top,
                Height = 62,
                Margin = new Padding(0, 0, 0, 10),
                BackColor = Green
            };
            update.Paint += PaintBorder;
            var updateIcon = new PictureBox
            {
                Image = EstimateUiIcons.Create(
                    EstimateUiIconKind.Refresh,
                    30,
                    Color.White),
                Location = new Point(14, 14),
                Size = new Size(35, 35),
                SizeMode = PictureBoxSizeMode.CenterImage
            };
            var updateTitle = new Label
            {
                Text = "Cập nhật THKP-TC",
                Location = new Point(58, 8),
                Height = 25,
                Font = new Font(
                    "Segoe UI",
                    10.5f,
                    FontStyle.Bold),
                ForeColor = Color.White,
                AutoEllipsis = true
            };
            var updateDetail = new Label
            {
                Text = "Tổng hợp số liệu từ các bảng dự toán",
                Location = new Point(59, 34),
                Height = 20,
                Font = new Font("Segoe UI", 7.8f),
                ForeColor = Color.White,
                AutoEllipsis = true
            };
            var updateArrow = new Label
            {
                Text = "›",
                Size = new Size(28, 42),
                Font = new Font("Segoe UI", 18f),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            update.Controls.Add(updateIcon);
            update.Controls.Add(updateTitle);
            update.Controls.Add(updateDetail);
            update.Controls.Add(updateArrow);
            update.Cursor = Cursors.Hand;
            EventHandler updateHandler =
                (s, e) => UpdateThkp();
            update.Click += updateHandler;
            updateIcon.Click += updateHandler;
            updateTitle.Click += updateHandler;
            updateDetail.Click += updateHandler;
            updateArrow.Click += updateHandler;
            update.Resize += (s, e) =>
            {
                updateTitle.Width =
                    Math.Max(130, update.ClientSize.Width - 100);
                updateDetail.Width =
                    Math.Max(130, update.ClientSize.Width - 100);
                updateArrow.Location =
                    new Point(update.ClientSize.Width - 36, 10);
            };
            content.Controls.Add(
                update,
                0,
                content.RowCount++);

            content.Controls.Add(
                BuildCheckHeader(),
                0,
                content.RowCount++);

            issuePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Margin = new Padding(0, 2, 0, 9),
                BackColor = Color.White
            };
            issuePanel.SizeChanged += (s, e) =>
            {
                foreach (Control control in issuePanel.Controls)
                    control.Width = Math.Max(
                        260,
                        issuePanel.ClientSize.Width - 2);
            };
            content.Controls.Add(
                issuePanel,
                0,
                content.RowCount++);

            content.Controls.Add(
                SectionTitle("Chức năng khác"),
                0,
                content.RowCount++);

            var actions = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 78,
                ColumnCount = 3,
                Margin = new Padding(0, 1, 0, 9)
            };
            for (int index = 0; index < 3; index++)
                actions.ColumnStyles.Add(
                    new ColumnStyle(SizeType.Percent, 33.333f));

            actions.Controls.Add(
                BuildActionCard(
                    "Kiểm tra hồ sơ",
                    "Kiểm tra dữ liệu, công thức",
                    EstimateUiIconKind.Document,
                    Green,
                    RefreshPreview),
                0,
                0);
            actions.Controls.Add(
                BuildActionCard(
                    "Xuất báo cáo",
                    "Mở THKP-TC để xuất/in",
                    EstimateUiIconKind.Document,
                    Blue,
                    ShowReportSheet),
                1,
                0);
            actions.Controls.Add(
                BuildActionCard(
                    "Mở thư mục hồ sơ",
                    "Mở thư mục chứa dự án",
                    EstimateUiIconKind.Folder,
                    Color.FromArgb(125, 74, 196),
                    OpenWorkbookFolder),
                2,
                0);
            content.Controls.Add(
                actions,
                0,
                content.RowCount++);

            statusLabel = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                MinimumSize = new Size(0, 56),
                Padding = new Padding(10, 9, 10, 9),
                BackColor = BlueSoft,
                ForeColor = Color.FromArgb(35, 88, 180),
                Font = new Font("Segoe UI", 8f),
                Text = "Bạn vẫn có thể mở module ngay cả khi chưa gắn định mức THKP-TC."
            };
            content.Controls.Add(
                statusLabel,
                0,
                content.RowCount++);

            RefreshPreview();
        }

        internal void RefreshPreview()
        {
            try
            {
                preview =
                    WorkbookEstimateV2CostLinkService
                        .BuildPreview(workbook);

                totalValue.Text =
                    preview.TotalWorkItemCount.ToString("N0");
                ratedValue.Text =
                    preview.RatedWorkItemCount.ToString("N0");
                warningValue.Text =
                    preview.WarningWorkItemCount.ToString("N0");
                formulaErrorValue.Text =
                    preview.FormulaErrorWorkItemCount.ToString("N0");

                RebuildIssues();

                if (preview.TotalWorkItemCount == 0)
                {
                    ShowStatus(
                        "Chưa đăng ký bảng công tác. " +
                        "Module vẫn mở bình thường; hãy đăng ký bảng ở bước Công tác.",
                        true);
                }
                else if (!preview.ThkpLinked)
                {
                    ShowStatus(
                        "Bấm Cập nhật THKP-TC để liên kết VL/NC/M/T từ các WorkItem. " +
                        "Các chi phí khác và thuế của mẫu THKP hiện hữu được giữ nguyên.",
                        false);
                }
                else
                {
                    ShowStatus(
                        "THKP-TC đang lấy chi phí trực tiếp bằng workbook Name/formula, " +
                        "không dùng số kết quả chết.",
                        false);
                }
            }
            catch (Exception ex)
            {
                totalValue.Text = "0";
                ratedValue.Text = "0";
                warningValue.Text = "!";
                formulaErrorValue.Text = "!";
                issuePanel.Controls.Clear();
                ShowStatus(
                    "Không đọc được trạng thái THKP-TC: " +
                    ex.Message,
                    true);
            }
        }

        private Control BuildHeader()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 66,
                BackColor = Color.White
            };
            var back = new Button
            {
                Text = "‹",
                Location = new Point(0, 2),
                Size = new Size(32, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = TextMuted,
                Font = new Font("Segoe UI", 15f),
                Cursor = Cursors.Hand
            };
            back.FlatAppearance.BorderSize = 0;
            back.Click += (s, e) =>
                backAction?.Invoke();

            var title = new Label
            {
                Text = "THKP-TC & Kiểm tra",
                Location = new Point(39, 1),
                AutoSize = true,
                Font = new Font(
                    "Segoe UI",
                    15f,
                    FontStyle.Bold),
                ForeColor = TextDark
            };
            var subtitle = new Label
            {
                Text = "Hỗ trợ lập, kiểm tra và tổng hợp THKP-TC",
                Location = new Point(40, 36),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = TextMuted
            };

            panel.Controls.Add(back);
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);
            return panel;
        }

        private static Control BuildCheckHeader()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 57,
                BackColor = Color.White
            };
            panel.Controls.Add(new PictureBox
            {
                Image = EstimateUiIcons.Create(
                    EstimateUiIconKind.Document,
                    29,
                    Green),
                Location = new Point(2, 10),
                Size = new Size(34, 34),
                SizeMode = PictureBoxSizeMode.CenterImage
            });
            panel.Controls.Add(new Label
            {
                Text = "Kết quả kiểm tra hồ sơ",
                Location = new Point(44, 6),
                AutoSize = true,
                Font = new Font(
                    "Segoe UI",
                    9.6f,
                    FontStyle.Bold),
                ForeColor = TextDark
            });
            panel.Controls.Add(new Label
            {
                Text = "Kiểm tra dữ liệu, đơn giá, công thức và tính nhất quán",
                Location = new Point(44, 29),
                AutoSize = true,
                Font = new Font("Segoe UI", 7.7f),
                ForeColor = TextMuted
            });
            return panel;
        }

        private void RebuildIssues()
        {
            foreach (Control control in
                issuePanel.Controls)
            {
                control.Dispose();
            }
            issuePanel.Controls.Clear();

            IReadOnlyList<EstimateV2CostIssue> items =
                preview?.Issues ??
                new EstimateV2CostIssue[0];

            foreach (EstimateV2CostIssue issue in
                items
                    .OrderByDescending(item =>
                        item.Severity)
                    .Take(5))
            {
                issuePanel.Controls.Add(
                    BuildIssueRow(issue));
            }
        }

        private Control BuildIssueRow(
            EstimateV2CostIssue issue)
        {
            Color color =
                issue.Severity ==
                    EstimateV2CostIssueSeverity.Error
                    ? Red
                    : issue.Severity ==
                        EstimateV2CostIssueSeverity.Warning
                        ? Amber
                        : Green;
            EstimateUiIconKind icon =
                issue.Severity ==
                    EstimateV2CostIssueSeverity.Error
                    ? EstimateUiIconKind.Error
                    : issue.Severity ==
                        EstimateV2CostIssueSeverity.Warning
                        ? EstimateUiIconKind.Warning
                        : EstimateUiIconKind.Check;

            var row = new Panel
            {
                Width = 380,
                Height = 58,
                Margin = new Padding(0, 0, 0, 4),
                BackColor = Color.White
            };
            row.Paint += PaintBorder;
            row.Controls.Add(new PictureBox
            {
                Image = EstimateUiIcons.Create(
                    icon,
                    27,
                    color),
                Location = new Point(9, 14),
                Size = new Size(32, 32),
                SizeMode = PictureBoxSizeMode.CenterImage
            });

            var title = new Label
            {
                Text = issue.Title,
                Location = new Point(49, 6),
                Height = 23,
                Font = new Font(
                    "Segoe UI",
                    8.4f,
                    FontStyle.Bold),
                ForeColor = color,
                AutoEllipsis = true
            };
            var detail = new Label
            {
                Text = issue.Detail,
                Location = new Point(49, 29),
                Height = 21,
                Font = new Font("Segoe UI", 7.2f),
                ForeColor = TextMuted,
                AutoEllipsis = true
            };
            var button = new Button
            {
                Text = "Xem chi tiết  ›",
                Size = new Size(88, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 7.3f),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top |
                    AnchorStyles.Right
            };
            button.FlatAppearance.BorderColor = Border;
            button.Click += (s, e) =>
                ShowIssue(issue);

            row.Controls.Add(title);
            row.Controls.Add(detail);
            row.Controls.Add(button);
            row.Resize += (s, e) =>
            {
                title.Width =
                    Math.Max(110, row.ClientSize.Width - 155);
                detail.Width =
                    Math.Max(110, row.ClientSize.Width - 155);
                button.Location =
                    new Point(row.ClientSize.Width - 97, 14);
            };
            return row;
        }

        private static Label AddMetric(
            TableLayoutPanel parent,
            int column,
            string title,
            EstimateUiIconKind iconKind,
            Color color)
        {
            var card = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 7, 0),
                BackColor = Color.White
            };
            card.Paint += PaintBorder;
            card.Controls.Add(new PictureBox
            {
                Image = EstimateUiIcons.Create(
                    iconKind,
                    23,
                    color),
                Location = new Point(8, 9),
                Size = new Size(28, 28),
                SizeMode = PictureBoxSizeMode.CenterImage
            });
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(38, 7),
                Height = 28,
                Font = new Font("Segoe UI", 7.1f),
                ForeColor = TextDark,
                AutoEllipsis = true
            };
            var value = new Label
            {
                Location = new Point(7, 39),
                Height = 38,
                Font = new Font(
                    "Segoe UI",
                    16f,
                    FontStyle.Bold),
                ForeColor = color,
                TextAlign = ContentAlignment.MiddleCenter
            };
            card.Controls.Add(titleLabel);
            card.Controls.Add(value);
            card.Resize += (s, e) =>
            {
                titleLabel.Width =
                    Math.Max(38, card.ClientSize.Width - 42);
                value.Width =
                    Math.Max(48, card.ClientSize.Width - 14);
            };
            parent.Controls.Add(card, column, 0);
            return value;
        }

        private static Label SectionTitle(
            string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Top,
                Height = 29,
                Margin = new Padding(0, 3, 0, 0),
                Font = new Font(
                    "Segoe UI",
                    9.3f,
                    FontStyle.Bold),
                ForeColor = TextDark,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private Control BuildActionCard(
            string title,
            string detail,
            EstimateUiIconKind iconKind,
            Color color,
            Action action)
        {
            var card = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 7, 0),
                BackColor = Color.White,
                Cursor = Cursors.Hand
            };
            card.Paint += PaintBorder;
            var icon = new PictureBox
            {
                Image = EstimateUiIcons.Create(
                    iconKind,
                    26,
                    color),
                Location = new Point(8, 10),
                Size = new Size(30, 30),
                SizeMode = PictureBoxSizeMode.CenterImage
            };
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(43, 7),
                Height = 23,
                Font = new Font(
                    "Segoe UI",
                    7.7f,
                    FontStyle.Bold),
                ForeColor = color,
                AutoEllipsis = true
            };
            var detailLabel = new Label
            {
                Text = detail,
                Location = new Point(8, 43),
                Height = 24,
                Font = new Font("Segoe UI", 6.8f),
                ForeColor = TextMuted,
                AutoEllipsis = true
            };
            EventHandler handler =
                (s, e) => action?.Invoke();
            card.Click += handler;
            icon.Click += handler;
            titleLabel.Click += handler;
            detailLabel.Click += handler;
            card.Controls.Add(icon);
            card.Controls.Add(titleLabel);
            card.Controls.Add(detailLabel);
            card.Resize += (s, e) =>
            {
                titleLabel.Width =
                    Math.Max(42, card.ClientSize.Width - 49);
                detailLabel.Width =
                    Math.Max(60, card.ClientSize.Width - 16);
            };
            return card;
        }

        private void UpdateThkp()
        {
            try
            {
                WorkbookEstimateV2CostWriteResult result =
                    WorkbookEstimateV2CostLinkWriter
                        .Apply(workbook);
                RefreshPreview();
                ShowStatus(
                    "Đã cập nhật " +
                    result.ThkpWorksheetName +
                    ": " +
                    result.LinkedWorkItemCount.ToString("N0") +
                    " công tác liên kết đơn giá." +
                    (result.MissingRateWorkItemCount > 0
                        ? " Còn " +
                            result.MissingRateWorkItemCount.ToString("N0") +
                            " công tác chưa đủ đơn giá."
                        : string.Empty),
                    result.MissingRateWorkItemCount > 0);
                ActivateSheet(
                    result.ThkpWorksheetName);
            }
            catch (Exception ex)
            {
                ShowStatus(
                    "Không cập nhật được THKP-TC: " +
                    ex.Message,
                    true);
            }
        }

        private void ShowIssue(
            EstimateV2CostIssue issue)
        {
            if (issue == null)
                return;

            if (string.Equals(
                issue.Code,
                "THKP",
                StringComparison.OrdinalIgnoreCase))
            {
                ActivateSheet("THKP-TC");
                return;
            }

            ActivateFirstRegisteredSource();
        }

        private void ShowReportSheet()
        {
            if (!ActivateSheet("THKP-TC"))
            {
                ShowStatus(
                    "Chưa có THKP-TC. Hãy cập nhật THKP-TC trước.",
                    true);
                return;
            }

            ShowStatus(
                "Đã mở THKP-TC. Chức năng xuất PDF/báo cáo hoàn chỉnh " +
                "được giữ cho task Báo cáo & Xuất in; V2-401 không tự tạo file ngoài.",
                false);
        }

        private void OpenWorkbookFolder()
        {
            try
            {
                string path = workbook.Path ?? string.Empty;
                if (path.Length == 0 ||
                    !Directory.Exists(path))
                {
                    ShowStatus(
                        "Workbook chưa được lưu vào thư mục trên máy.",
                        true);
                    return;
                }

                Process.Start(
                    new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
            }
            catch (Exception ex)
            {
                ShowStatus(
                    "Không mở được thư mục hồ sơ: " +
                    ex.Message,
                    true);
            }
        }

        private void ActivateFirstRegisteredSource()
        {
            try
            {
                EstimateV2RegisteredSource source =
                    WorkbookEstimateV2RegistrationService
                        .ListRegistered(workbook)
                        .FirstOrDefault();
                if (source == null)
                    return;
                ActivateSheet(source.WorksheetName);
            }
            catch
            {
            }
        }

        private bool ActivateSheet(string name)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1;
                    index <= sheets.Count;
                    index++)
                {
                    Excel.Worksheet candidate = null;
                    try
                    {
                        candidate = sheets.Item[index]
                            as Excel.Worksheet;
                        if (candidate != null &&
                            string.Equals(
                                candidate.Name,
                                name,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            sheet = candidate;
                            candidate = null;
                            break;
                        }
                    }
                    finally
                    {
                        Release(candidate);
                    }
                }

                if (sheet == null)
                    return false;
                sheet.Activate();
                return true;
            }
            finally
            {
                Release(sheet);
                Release(sheets);
            }
        }

        private void ShowStatus(
            string text,
            bool warning)
        {
            statusLabel.BackColor =
                warning ? AmberSoft : BlueSoft;
            statusLabel.ForeColor =
                warning
                    ? Color.DarkGoldenrod
                    : Color.FromArgb(35, 88, 180);
            statusLabel.Text =
                text ?? string.Empty;
        }

        private static void PaintBorder(
            object sender,
            PaintEventArgs e)
        {
            Control control = (Control)sender;
            using (var pen = new Pen(Border))
            {
                e.Graphics.DrawRectangle(
                    pen,
                    0,
                    0,
                    Math.Max(0, control.Width - 1),
                    Math.Max(0, control.Height - 1));
            }
        }

        private static void Release(
            object value)
        {
            if (value != null &&
                Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
