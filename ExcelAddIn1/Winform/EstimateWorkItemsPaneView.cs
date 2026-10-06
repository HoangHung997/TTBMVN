using ExcelAddIn1.Funtion;
using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal sealed class EstimateWorkItemsPaneView : UserControl
    {
        private static readonly Color Green = Color.FromArgb(0, 137, 70);
        private static readonly Color GreenDark = Color.FromArgb(0, 103, 55);
        private static readonly Color GreenSoft = Color.FromArgb(235, 248, 239);
        private static readonly Color Border = Color.FromArgb(222, 228, 223);
        private static readonly Color TextDark = Color.FromArgb(33, 43, 54);
        private static readonly Color TextMuted = Color.FromArgb(92, 103, 112);
        private static readonly Color Amber = Color.FromArgb(244, 166, 35);

        private readonly Excel.Workbook workbook;
        private readonly Action backAction;
        private readonly Action nextAction;
        private readonly Label totalValue;
        private readonly Label registeredValue;
        private readonly Label unboundValue;
        private readonly FlowLayoutPanel sourceList;
        private readonly Label mapWorkCode;
        private readonly Label mapNorm;
        private readonly Label mapDescription;
        private readonly Label mapUnit;
        private readonly Label mapQuantity;
        private readonly Label statusLabel;

        internal EstimateWorkItemsPaneView(
            Excel.Workbook workbook,
            Action backAction,
            Action nextAction)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.backAction = backAction;
            this.nextAction = nextAction;
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;

            var root = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            Controls.Add(root);
            var content = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(14, 10, 14, 14),
                BackColor = Color.White,
                ColumnCount = 1
            };
            root.Controls.Add(content);

            content.Controls.Add(BuildHeader(), 0, content.RowCount++);

            var metrics = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 98,
                ColumnCount = 3,
                Margin = new Padding(0, 5, 0, 10)
            };
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34f));
            totalValue = AddMetric(metrics, 0, "Tổng công tác", EstimateUiIconKind.Clipboard, Green);
            registeredValue = AddMetric(metrics, 1, "Đã nhận diện", EstimateUiIconKind.Check, Green);
            unboundValue = AddMetric(metrics, 2, "Chưa gắn", EstimateUiIconKind.Warning, Amber);
            content.Controls.Add(metrics, 0, content.RowCount++);

            content.Controls.Add(SectionTitle("1", "Chọn bảng công tác",
                "Hệ thống đã phát hiện bảng công tác trong workbook hiện tại."), 0, content.RowCount++);

            sourceList = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Margin = new Padding(0, 2, 0, 7),
                BackColor = Color.White
            };
            sourceList.SizeChanged += (s, e) =>
            {
                foreach (Control control in sourceList.Controls)
                    control.Width = Math.Max(260, sourceList.ClientSize.Width - 2);
            };
            content.Controls.Add(sourceList, 0, content.RowCount++);

            content.Controls.Add(SectionTitle("2", "Nhận diện cột dữ liệu",
                "Xác định các cột trong bảng công tác của bạn."), 0, content.RowCount++);

            var mapCard = new Panel
            {
                Dock = DockStyle.Top,
                Height = 164,
                BackColor = Color.White,
                Margin = new Padding(0, 2, 0, 7),
                Padding = new Padding(8)
            };
            mapCard.Paint += PaintBorder;
            var map = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 5
            };
            map.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
            map.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 5; i++)
                map.RowStyles.Add(new RowStyle(SizeType.Percent, 20));
            mapWorkCode = AddMapRow(map, 0, "Mã công tác");
            mapNorm = AddMapRow(map, 1, "Định mức");
            mapDescription = AddMapRow(map, 2, "Mô tả công việc");
            mapUnit = AddMapRow(map, 3, "Đơn vị");
            mapQuantity = AddMapRow(map, 4, "Khối lượng");
            mapCard.Controls.Add(map);
            content.Controls.Add(mapCard, 0, content.RowCount++);

            var commands = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 46,
                ColumnCount = 3,
                Margin = new Padding(0, 0, 0, 7)
            };
            for (int i = 0; i < 3; i++)
                commands.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            var scan = SecondaryButton("Quét vùng chọn", Color.FromArgb(52, 103, 235));
            var recognize = SecondaryButton("Nhận diện lại", Color.FromArgb(74, 85, 94));
            var register = SecondaryButton("Đăng ký bảng", Color.FromArgb(52, 103, 235));
            scan.Click += (s, e) => RegisterCurrentSelection();
            recognize.Click += (s, e) => ReconcileRegistered();
            register.Click += (s, e) => RegisterCurrentSelection();
            commands.Controls.Add(scan, 0, 0);
            commands.Controls.Add(recognize, 1, 0);
            commands.Controls.Add(register, 2, 0);
            content.Controls.Add(commands, 0, content.RowCount++);

            var next = new Button
            {
                Text = "Tiếp tục gắn định mức   →",
                Dock = DockStyle.Top,
                Height = 43,
                Margin = new Padding(0, 0, 0, 8),
                BackColor = Green,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            next.FlatAppearance.BorderSize = 0;
            next.Click += (s, e) => this.nextAction?.Invoke();
            content.Controls.Add(next, 0, content.RowCount++);

            statusLabel = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                MinimumSize = new Size(0, 55),
                Padding = new Padding(10, 9, 10, 9),
                BackColor = GreenSoft,
                ForeColor = GreenDark,
                Font = new Font("Segoe UI", 8.1f),
                Text = "Có thể mở và sử dụng module này ngay cả khi chưa gắn THKP-TC."
            };
            content.Controls.Add(statusLabel, 0, content.RowCount++);

            RefreshView();
        }

        internal void RefreshView()
        {
            try
            {
                EstimateV2ReconcileResult reconcile =
                    WorkbookEstimateV2RegistrationService.ReconcileAll(workbook);
                var sources = WorkbookEstimateV2RegistrationService.ListRegistered(workbook);
                PopulateSources(sources);

                totalValue.Text = reconcile.WorkItemCount.ToString("N0");
                registeredValue.Text = reconcile.WorkItemCount.ToString("N0");

                ExcelAddIn1.Core.EstimateV2State state;
                int unbound = 0;
                if (WorkbookEstimateV2StateService.TryLoad(workbook, out state))
                    unbound = state.WorkItems.Count(item => !item.IsOrphaned && !item.HasNormBinding);
                unboundValue.Text = unbound.ToString("N0");

                if (sources.Count > 0)
                {
                    ApplySourceMap(sources[0]);
                    statusLabel.Text = "Đã nhớ bảng công tác và metadata ẩn. Chèn/xóa/sắp xếp dòng không dùng RowIndex làm khóa.";
                }
                else
                {
                    ClearSourceMap();
                    statusLabel.Text = "Chưa đăng ký bảng công tác. Hãy chọn vùng gồm dòng tiêu đề và các dòng công tác rồi bấm Quét vùng chọn.";
                }
            }
            catch (Exception ex)
            {
                statusLabel.Text = "Không đọc được bảng công tác: " + ex.Message;
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.BackColor = Color.MistyRose;
            }
        }

        private void RegisterCurrentSelection()
        {
            Excel.Range range = null;
            try
            {
                range = workbook.Application.Selection as Excel.Range;
                if (range == null)
                    throw new InvalidOperationException("Hãy chọn vùng bảng công tác trên Excel trước.");
                EstimateV2RegistrationResult result =
                    WorkbookEstimateV2RegistrationService.RegisterSelectedRange(
                        workbook,
                        range);
                ApplySourceMap(result.Source);
                statusLabel.ForeColor = GreenDark;
                statusLabel.BackColor = GreenSoft;
                statusLabel.Text = "Đã đăng ký " + result.Reconcile.WorkItemCount +
                    " công tác; tạo mới " + result.Reconcile.CreatedCount +
                    " WorkItemId, phục hồi " + result.Reconcile.RecoveredCount + ".";
                RefreshView();
            }
            catch (Exception ex)
            {
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.BackColor = Color.MistyRose;
                statusLabel.Text = ex.Message;
            }
            finally
            {
                Release(range);
            }
        }

        private void ReconcileRegistered()
        {
            try
            {
                EstimateV2ReconcileResult result =
                    WorkbookEstimateV2RegistrationService.ReconcileAll(workbook);
                statusLabel.ForeColor = GreenDark;
                statusLabel.BackColor = GreenSoft;
                statusLabel.Text = "Đã nhận diện lại: " + result.WorkItemCount +
                    " công tác, phục hồi " + result.RecoveredCount +
                    ", ID trùng xử lý " + result.DuplicateIdCount + ".";
                RefreshView();
            }
            catch (Exception ex)
            {
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.BackColor = Color.MistyRose;
                statusLabel.Text = ex.Message;
            }
        }

        private void PopulateSources(
            System.Collections.Generic.IReadOnlyList<EstimateV2RegisteredSource> sources)
        {
            foreach (Control control in sourceList.Controls)
                control.Dispose();
            sourceList.Controls.Clear();

            if (sources.Count == 0)
            {
                sourceList.Controls.Add(new Label
                {
                    Height = 38,
                    Width = 340,
                    Text = "○ Chưa có bảng công tác được đăng ký",
                    ForeColor = TextMuted,
                    Font = new Font("Segoe UI", 8.4f),
                    TextAlign = ContentAlignment.MiddleLeft
                });
                return;
            }

            int number = 1;
            foreach (EstimateV2RegisteredSource source in sources)
            {
                var row = new Panel
                {
                    Height = 43,
                    Width = 360,
                    Margin = new Padding(0, 0, 0, 4),
                    BackColor = Color.White
                };
                row.Paint += PaintBorder;
                var radio = new Label
                {
                    Text = number == 1 ? "●" : "○",
                    ForeColor = number == 1 ? Color.FromArgb(52, 103, 235) : TextMuted,
                    Location = new Point(8, 9),
                    Size = new Size(26, 24),
                    Font = new Font("Segoe UI", 11f),
                    TextAlign = ContentAlignment.MiddleCenter
                };
                var name = new Label
                {
                    Text = "Bảng công tác " + number + " (" + source.WorksheetName + ")",
                    Location = new Point(38, 5),
                    Height = 22,
                    Font = new Font("Segoe UI", 8.2f, FontStyle.Bold),
                    ForeColor = TextDark,
                    AutoEllipsis = true
                };
                var rows = new Label
                {
                    Text = Math.Max(0, source.LastDataRow - source.FirstDataRow + 1) + " dòng",
                    Height = 22,
                    TextAlign = ContentAlignment.MiddleRight,
                    Font = new Font("Segoe UI", 7.6f),
                    ForeColor = TextMuted,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right
                };
                row.Controls.Add(radio);
                row.Controls.Add(name);
                row.Controls.Add(rows);
                row.Resize += (s, e) =>
                {
                    name.Width = Math.Max(120, row.ClientSize.Width - 145);
                    rows.Location = new Point(row.ClientSize.Width - 92, 6);
                    rows.Width = 82;
                };
                sourceList.Controls.Add(row);
                number++;
            }
        }

        private void ApplySourceMap(EstimateV2RegisteredSource source)
        {
            mapWorkCode.Text = ExcelAddIn1.Core.ExcelColumnAddress.ToLetters(source.Columns.WorkCodeColumn) + ": Mã công tác";
            mapNorm.Text = ExcelAddIn1.Core.ExcelColumnAddress.ToLetters(source.Columns.NormDisplayColumn) + ": Định mức";
            mapDescription.Text = ExcelAddIn1.Core.ExcelColumnAddress.ToLetters(source.Columns.DescriptionColumn) + ": Mô tả công việc";
            mapUnit.Text = ExcelAddIn1.Core.ExcelColumnAddress.ToLetters(source.Columns.UnitColumn) + ": Đơn vị";
            mapQuantity.Text = ExcelAddIn1.Core.ExcelColumnAddress.ToLetters(source.Columns.QuantityColumn) + ": Khối lượng";
        }

        private void ClearSourceMap()
        {
            mapWorkCode.Text = "(chưa nhận diện)";
            mapNorm.Text = "(chưa nhận diện)";
            mapDescription.Text = "(chưa nhận diện)";
            mapUnit.Text = "(chưa nhận diện)";
            mapQuantity.Text = "(chưa nhận diện)";
        }

        private Control BuildHeader()
        {
            var panel = new Panel { Dock = DockStyle.Top, Height = 77, BackColor = Color.White };
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
            back.Click += (s, e) => backAction?.Invoke();
            var overline = new Label
            {
                Text = "Trợ lý Dự toán",
                Location = new Point(39, 2),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = TextDark
            };
            var title = new Label
            {
                Text = "Công tác",
                Location = new Point(39, 23),
                AutoSize = true,
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor = TextDark
            };
            var subtitle = new Label
            {
                Text = "Nhận diện bảng công tác và gắn định mức",
                Location = new Point(40, 53),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.4f),
                ForeColor = TextMuted
            };
            panel.Controls.Add(back);
            panel.Controls.Add(overline);
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);
            return panel;
        }

        private static Control SectionTitle(string number, string title, string subtitle)
        {
            var panel = new Panel { Dock = DockStyle.Top, Height = 57, BackColor = Color.White };
            var circle = new Label
            {
                Text = number,
                Location = new Point(0, 7),
                Size = new Size(34, 34),
                BackColor = Green,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(43, 3),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.3f, FontStyle.Bold),
                ForeColor = GreenDark
            };
            var subtitleLabel = new Label
            {
                Text = subtitle,
                Location = new Point(43, 26),
                Height = 27,
                Width = 320,
                Font = new Font("Segoe UI", 7.8f),
                ForeColor = TextMuted,
                AutoEllipsis = true
            };
            panel.Controls.Add(circle);
            panel.Controls.Add(titleLabel);
            panel.Controls.Add(subtitleLabel);
            return panel;
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
            var icon = new PictureBox
            {
                Image = EstimateUiIcons.Create(iconKind, 23, color),
                Location = new Point(9, 10),
                Size = new Size(28, 28),
                SizeMode = PictureBoxSizeMode.CenterImage
            };
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(40, 9),
                Height = 24,
                Font = new Font("Segoe UI", 7.4f),
                ForeColor = TextDark,
                AutoEllipsis = true
            };
            var value = new Label
            {
                Location = new Point(8, 39),
                Height = 38,
                Font = new Font("Segoe UI", 16f, FontStyle.Bold),
                ForeColor = color,
                TextAlign = ContentAlignment.MiddleCenter
            };
            card.Controls.Add(icon);
            card.Controls.Add(titleLabel);
            card.Controls.Add(value);
            card.Resize += (s, e) =>
            {
                titleLabel.Width = Math.Max(45, card.ClientSize.Width - 44);
                value.Width = Math.Max(50, card.ClientSize.Width - 16);
            };
            parent.Controls.Add(card, column, 0);
            return value;
        }

        private static Label AddMapRow(TableLayoutPanel map, int row, string title)
        {
            map.Controls.Add(new Label
            {
                Text = title,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.1f),
                ForeColor = TextDark,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, row);
            var value = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.1f),
                ForeColor = TextDark,
                BackColor = Color.FromArgb(249, 250, 249),
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(7, 0, 4, 0)
            };
            map.Controls.Add(value, 1, row);
            return value;
        }

        private static Button SecondaryButton(string text, Color color)
        {
            var button = new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 6, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = color,
                Font = new Font("Segoe UI", 7.7f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(205, 213, 220);
            button.FlatAppearance.BorderSize = 1;
            return button;
        }

        private static void PaintBorder(object sender, PaintEventArgs e)
        {
            Control control = (Control)sender;
            using (var pen = new Pen(Border))
                e.Graphics.DrawRectangle(pen, 0, 0, Math.Max(0, control.Width - 1), Math.Max(0, control.Height - 1));
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
