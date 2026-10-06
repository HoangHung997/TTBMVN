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
    internal sealed class EstimateResourcesPaneView : UserControl
    {
        private static readonly Color Green = Color.FromArgb(0, 137, 70);
        private static readonly Color GreenDark = Color.FromArgb(0, 103, 55);
        private static readonly Color GreenSoft = Color.FromArgb(235, 248, 239);
        private static readonly Color Blue = Color.FromArgb(45, 112, 229);
        private static readonly Color BlueSoft = Color.FromArgb(237, 244, 255);
        private static readonly Color Amber = Color.FromArgb(244, 166, 35);
        private static readonly Color AmberSoft = Color.FromArgb(255, 248, 226);
        private static readonly Color Red = Color.FromArgb(220, 53, 69);
        private static readonly Color RedSoft = Color.FromArgb(255, 238, 239);
        private static readonly Color Border = Color.FromArgb(222, 228, 223);
        private static readonly Color TextDark = Color.FromArgb(33, 43, 54);
        private static readonly Color TextMuted = Color.FromArgb(92, 103, 112);

        private readonly Excel.Workbook workbook;
        private readonly Action backAction;
        private readonly Label materialValue;
        private readonly Label laborValue;
        private readonly Label machineValue;
        private readonly Label missingValue;
        private readonly Label materialSummary;
        private readonly Label laborSummary;
        private readonly Label machineSummary;
        private readonly DataGridView missingGrid;
        private readonly Label statusLabel;
        private WorkbookEstimateV2ResourcePreview preview;
        private PriceProfile previewProfile;

        internal EstimateResourcesPaneView(
            Excel.Workbook workbook,
            Action backAction)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
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
                ColumnCount = 4,
                Margin = new Padding(0, 5, 0, 10)
            };
            for (int i = 0; i < 4; i++)
                metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            materialValue = AddMetric(
                metrics, 0, "Số vật liệu", EstimateUiIconKind.Clipboard, Green);
            laborValue = AddMetric(
                metrics, 1, "Số nhân công", EstimateUiIconKind.Document, Blue);
            machineValue = AddMetric(
                metrics, 2, "Số máy", EstimateUiIconKind.Settings, Amber);
            missingValue = AddMetric(
                metrics, 3, "Thiếu giá", EstimateUiIconKind.Warning, Red);
            content.Controls.Add(metrics, 0, content.RowCount++);

            content.Controls.Add(SectionTitle("Tổng quan theo nhóm"), 0, content.RowCount++);

            var groups = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 112,
                ColumnCount = 3,
                Margin = new Padding(0, 3, 0, 9)
            };
            groups.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            groups.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            groups.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34f));
            materialSummary = AddGroupCard(
                groups, 0, "Vật liệu", EstimateUiIconKind.Clipboard, Green, GreenSoft);
            laborSummary = AddGroupCard(
                groups, 1, "Nhân công", EstimateUiIconKind.Document, Blue, BlueSoft);
            machineSummary = AddGroupCard(
                groups, 2, "Máy thi công", EstimateUiIconKind.Settings, Amber, AmberSoft);
            content.Controls.Add(groups, 0, content.RowCount++);

            var missingTitle = new Label
            {
                Text = "Danh sách khoản mục chưa có giá",
                Dock = DockStyle.Top,
                Height = 30,
                Margin = new Padding(0, 0, 0, 2),
                ForeColor = Red,
                Font = new Font("Segoe UI", 9.2f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            content.Controls.Add(missingTitle, 0, content.RowCount++);

            var missingCard = new Panel
            {
                Dock = DockStyle.Top,
                Height = 168,
                Margin = new Padding(0, 0, 0, 9),
                BackColor = RedSoft,
                Padding = new Padding(6)
            };
            missingCard.Paint += PaintBorder;
            missingGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                ReadOnly = true,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Segoe UI", 7.7f)
            };
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Mã hiệu",
                Width = 91
            });
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Tên tài nguyên",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 120
            });
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Đơn vị",
                Width = 56
            });
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Nhóm",
                Width = 76
            });
            missingCard.Controls.Add(missingGrid);
            content.Controls.Add(missingCard, 0, content.RowCount++);

            content.Controls.Add(SectionTitle("Chức năng chính"), 0, content.RowCount++);

            var actions = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 112,
                ColumnCount = 2,
                RowCount = 2,
                Margin = new Padding(0, 3, 0, 9)
            };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            actions.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            actions.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            Button aggregate = FeatureButton(
                "Tổng hợp tài nguyên",
                "Từ các công tác, định mức",
                EstimateUiIconKind.Clipboard,
                Green,
                GreenSoft);
            Button price = FeatureButton(
                "Cập nhật giá",
                "Mở/đồng bộ bảng giá",
                EstimateUiIconKind.Refresh,
                Blue,
                BlueSoft);
            Button labor = FeatureButton(
                "Sinh công thức NC",
                "Tạo công thức nhân công",
                EstimateUiIconKind.Link,
                Color.FromArgb(110, 73, 200),
                Color.FromArgb(245, 240, 255));
            Button machine = FeatureButton(
                "Sinh công thức giá ca máy",
                "Tạo công thức ca máy",
                EstimateUiIconKind.Settings,
                Color.FromArgb(230, 132, 28),
                AmberSoft);

            aggregate.Click += (s, e) => RefreshPreview();
            price.Click += (s, e) => ActivateResourceSheet();
            labor.Click += (s, e) => ShowPendingWriter(
                "Công thức nhân công sẽ được sinh trực tiếp trong sheet VL-NC-M; không ghi giá ngày công chết.");
            machine.Click += (s, e) => ShowPendingWriter(
                "Giá ca máy sẽ được sinh từ định mức máy và link giá nhiên liệu/nhân công trong VL-NC-M.");

            actions.Controls.Add(aggregate, 0, 0);
            actions.Controls.Add(price, 1, 0);
            actions.Controls.Add(labor, 0, 1);
            actions.Controls.Add(machine, 1, 1);
            content.Controls.Add(actions, 0, content.RowCount++);

            statusLabel = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                MinimumSize = new Size(0, 62),
                Padding = new Padding(10, 9, 10, 9),
                BackColor = BlueSoft,
                ForeColor = Color.FromArgb(35, 88, 180),
                Font = new Font("Segoe UI", 8.1f),
                Text = "Giá trị trong bảng này phải được sinh từ công thức hoặc liên kết; không có kết quả tính toán số chết."
            };
            content.Controls.Add(statusLabel, 0, content.RowCount++);

            RefreshPreview();
        }

        internal void RefreshPreview()
        {
            try
            {
                preview = WorkbookEstimateV2ResourceService.BuildPreview(workbook);
                previewProfile = SelectPreviewProfile();

                materialValue.Text = preview.Plan.Materials.Count.ToString("N0");
                laborValue.Text = preview.Plan.Labor.Count.ToString("N0");
                machineValue.Text = preview.Plan.Machines.Count.ToString("N0");

                var missing = BuildMissingItems();
                missingValue.Text = missing.Count.ToString("N0");
                PopulateMissing(missing);

                materialSummary.Text = FormatGroupSummary(
                    preview.Plan.Materials,
                    NormResourceKind.Material);
                laborSummary.Text = FormatGroupSummary(
                    preview.Plan.Labor,
                    NormResourceKind.Labor);
                machineSummary.Text = FormatGroupSummary(
                    preview.Plan.Machines,
                    NormResourceKind.Machine);

                if (preview.MissingPackageBindings.Count > 0)
                {
                    ShowStatus(
                        "Có binding chưa resolve được package: " +
                        string.Join(" | ", preview.MissingPackageBindings.Take(2)) +
                        (preview.MissingPackageBindings.Count > 2 ? " ..." : string.Empty),
                        true);
                }
                else if (preview.UnresolvedLogicalResources.Count > 0)
                {
                    ShowStatus(
                        "Có tài nguyên logic cần ràng buộc tài nguyên thật trước khi sinh đơn giá: " +
                        string.Join(", ", preview.UnresolvedLogicalResources),
                        true);
                }
                else if (preview.Plan.BoundWorkItemCount == 0)
                {
                    ShowStatus(
                        "Chưa có công tác được gắn định mức. Hãy hoàn thành bước Gắn định mức trước.",
                        true);
                }
                else
                {
                    ShowStatus(
                        "Đã gom " + preview.Plan.Resources.Count +
                        " tài nguyên unique từ " + preview.Plan.BoundWorkItemCount +
                        " công tác và " + preview.Plan.UniqueNormBindingCount +
                        " định mức/variant.", false);
                }
            }
            catch (Exception ex)
            {
                materialValue.Text = "0";
                laborValue.Text = "0";
                machineValue.Text = "0";
                missingValue.Text = "!";
                missingGrid.Rows.Clear();
                ShowStatus("Không tổng hợp được VL-NC-M: " + ex.Message, true);
            }
        }

        private List<MissingItem> BuildMissingItems()
        {
            var result = new List<MissingItem>();
            if (preview == null)
                return result;

            foreach (EstimateV2ResourceRequirement item in preview.Plan.Resources)
            {
                if (!item.RequiresUnitPrice)
                    continue;

                PriceProfileEntry entry;
                bool hasPrice = previewProfile != null &&
                    previewProfile.TryFind(item.Code, out entry);
                if (hasPrice)
                    continue;

                result.Add(new MissingItem(
                    item.Code,
                    ResolveName(item.Code),
                    item.Unit,
                    KindText(item.Kind)));
            }

            foreach (string logical in preview.UnresolvedLogicalResources)
            {
                if (result.Any(item => string.Equals(
                    item.Code, logical, StringComparison.OrdinalIgnoreCase)))
                    continue;
                result.Add(new MissingItem(
                    logical,
                    "Tài nguyên logic chưa ràng buộc",
                    string.Empty,
                    "Máy"));
            }

            return result
                .OrderBy(item => item.Group, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Code, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private PriceProfile SelectPreviewProfile()
        {
            PriceProfilePortfolio portfolio;
            if (!WorkbookPriceProfilePortfolioService.TryLoad(workbook, out portfolio) ||
                portfolio.Profiles.Count == 0)
                return null;

            PriceProfile nonState;
            if (portfolio.TryGet(MachineRateAudience.NonStateSalary, out nonState))
                return nonState;
            return portfolio.Profiles[0];
        }

        private string ResolveName(string code)
        {
            PriceProfileEntry entry;
            if (previewProfile != null && previewProfile.TryFind(code, out entry))
                return entry.DisplayName;
            return code;
        }

        private string FormatGroupSummary(
            IReadOnlyList<EstimateV2ResourceRequirement> items,
            NormResourceKind kind)
        {
            decimal total = 0m;
            int priced = 0;
            foreach (EstimateV2ResourceRequirement item in items)
            {
                PriceProfileEntry entry;
                if (previewProfile != null &&
                    item.RequiresUnitPrice &&
                    previewProfile.TryFind(item.Code, out entry))
                {
                    total += entry.AppliedUnitPriceVnd;
                    priced++;
                }
            }

            if (previewProfile == null)
                return items.Count.ToString("N0") + " khoản mục\r\nChưa có hồ sơ giá";
            return items.Count.ToString("N0") + " khoản mục\r\n" +
                priced.ToString("N0") + " có giá" +
                (total > 0m
                    ? " • " + total.ToString("#,##0", CultureInfo.CurrentCulture) + " đ"
                    : string.Empty);
        }

        private void PopulateMissing(IEnumerable<MissingItem> items)
        {
            missingGrid.Rows.Clear();
            foreach (MissingItem item in items)
                missingGrid.Rows.Add(item.Code, item.Name, item.Unit, item.Group);
        }

        private void ActivateResourceSheet()
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet worksheet = null;
                    try
                    {
                        worksheet = sheets.Item[index] as Excel.Worksheet;
                        if (worksheet != null &&
                            string.Equals(
                                worksheet.Name,
                                "VL-NC-M",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            worksheet.Activate();
                            ShowStatus(
                                "Đã mở sheet VL-NC-M. Writer V2 sẽ giữ A:F là vùng in và đặt metadata ngoài vùng in.",
                                false);
                            return;
                        }
                    }
                    finally
                    {
                        Release(worksheet);
                    }
                }

                ShowStatus(
                    "Workbook chưa có sheet VL-NC-M. Task writer sẽ tạo sheet theo đúng mẫu in khi được bật.",
                    true);
            }
            finally
            {
                Release(sheets);
            }
        }

        private void ShowPendingWriter(string message)
        {
            ShowStatus(
                message + " Backend writer đang được triển khai trong V2-201.",
                false);
        }

        private void ShowStatus(string message, bool warning)
        {
            statusLabel.BackColor = warning ? AmberSoft : BlueSoft;
            statusLabel.ForeColor = warning ? Color.DarkGoldenrod : Color.FromArgb(35, 88, 180);
            statusLabel.Text = message;
        }

        private Control BuildHeader()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 76,
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
            back.Click += (s, e) => backAction?.Invoke();

            var title = new Label
            {
                Text = "VL-NC-M",
                Location = new Point(39, 5),
                AutoSize = true,
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor = TextDark
            };
            var subtitle = new Label
            {
                Text = "Giá vật liệu - nhân công - máy thi công",
                Location = new Point(40, 38),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.4f),
                ForeColor = TextMuted
            };
            panel.Controls.Add(back);
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);
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
            card.Controls.Add(new PictureBox
            {
                Image = EstimateUiIcons.Create(iconKind, 23, color),
                Location = new Point(8, 10),
                Size = new Size(28, 28),
                SizeMode = PictureBoxSizeMode.CenterImage
            });
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(38, 8),
                Height = 25,
                Font = new Font("Segoe UI", 7.2f),
                ForeColor = TextDark,
                AutoEllipsis = true
            };
            var value = new Label
            {
                Location = new Point(7, 39),
                Height = 38,
                Font = new Font("Segoe UI", 16f, FontStyle.Bold),
                ForeColor = color,
                TextAlign = ContentAlignment.MiddleCenter
            };
            card.Controls.Add(titleLabel);
            card.Controls.Add(value);
            card.Resize += (s, e) =>
            {
                titleLabel.Width = Math.Max(38, card.ClientSize.Width - 42);
                value.Width = Math.Max(48, card.ClientSize.Width - 14);
            };
            parent.Controls.Add(card, column, 0);
            return value;
        }

        private static Label AddGroupCard(
            TableLayoutPanel parent,
            int column,
            string title,
            EstimateUiIconKind iconKind,
            Color color,
            Color background)
        {
            var card = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 7, 0),
                BackColor = background,
                Padding = new Padding(9)
            };
            card.Paint += PaintBorder;
            card.Controls.Add(new PictureBox
            {
                Image = EstimateUiIcons.Create(iconKind, 25, color),
                Location = new Point(10, 10),
                Size = new Size(30, 30),
                SizeMode = PictureBoxSizeMode.CenterImage
            });
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(45, 10),
                Height = 25,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = color,
                AutoEllipsis = true
            };
            var summary = new Label
            {
                Location = new Point(10, 47),
                Height = 52,
                Font = new Font("Segoe UI", 8f),
                ForeColor = TextDark,
                TextAlign = ContentAlignment.MiddleCenter
            };
            card.Controls.Add(titleLabel);
            card.Controls.Add(summary);
            card.Resize += (s, e) =>
            {
                titleLabel.Width = Math.Max(50, card.ClientSize.Width - 52);
                summary.Width = Math.Max(60, card.ClientSize.Width - 20);
            };
            parent.Controls.Add(card, column, 0);
            return summary;
        }

        private static Label SectionTitle(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Top,
                Height = 29,
                Margin = new Padding(0, 3, 0, 0),
                Font = new Font("Segoe UI", 9.3f, FontStyle.Bold),
                ForeColor = TextDark,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static Button FeatureButton(
            string title,
            string subtitle,
            EstimateUiIconKind iconKind,
            Color color,
            Color background)
        {
            var button = new Button
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 7, 6),
                FlatStyle = FlatStyle.Flat,
                BackColor = background,
                ForeColor = color,
                Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                Text = title + "\r\n" + subtitle,
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                Image = EstimateUiIcons.Create(iconKind, 22, color),
                ImageAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText
            };
            button.FlatAppearance.BorderColor = Border;
            button.FlatAppearance.BorderSize = 1;
            return button;
        }

        private static string KindText(NormResourceKind kind)
        {
            switch (kind)
            {
                case NormResourceKind.Material: return "Vật liệu";
                case NormResourceKind.Labor: return "Nhân công";
                case NormResourceKind.Machine: return "Máy";
                default: return kind.ToString();
            }
        }

        private static void PaintBorder(object sender, PaintEventArgs e)
        {
            Control control = (Control)sender;
            using (var pen = new Pen(Border))
                e.Graphics.DrawRectangle(
                    pen,
                    0,
                    0,
                    Math.Max(0, control.Width - 1),
                    Math.Max(0, control.Height - 1));
        }

        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value))
                System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
        }

        private sealed class MissingItem
        {
            internal MissingItem(string code, string name, string unit, string group)
            {
                Code = code ?? string.Empty;
                Name = name ?? string.Empty;
                Unit = unit ?? string.Empty;
                Group = group ?? string.Empty;
            }

            internal string Code { get; }
            internal string Name { get; }
            internal string Unit { get; }
            internal string Group { get; }
        }
    }
}
