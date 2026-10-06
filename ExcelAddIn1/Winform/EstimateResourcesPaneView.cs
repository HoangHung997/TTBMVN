using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
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
        private readonly GroupSummaryView materialSummary;
        private readonly GroupSummaryView laborSummary;
        private readonly GroupSummaryView machineSummary;
        private readonly DataGridView missingGrid;
        private readonly Label missingTitleLabel;
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
                metrics, 0, "Số vật liệu", EstimateUiIconKind.Materials, Green);
            laborValue = AddMetric(
                metrics, 1, "Số nhân công", EstimateUiIconKind.Worker, Blue);
            machineValue = AddMetric(
                metrics, 2, "Số máy", EstimateUiIconKind.Excavator, Amber);
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
                groups, 0, "Vật liệu", EstimateUiIconKind.Materials, Green, GreenSoft);
            laborSummary = AddGroupCard(
                groups, 1, "Nhân công", EstimateUiIconKind.Worker, Blue, BlueSoft);
            machineSummary = AddGroupCard(
                groups, 2, "Máy thi công", EstimateUiIconKind.Excavator, Amber, AmberSoft);
            content.Controls.Add(groups, 0, content.RowCount++);

            var missingCard = new Panel
            {
                Dock = DockStyle.Top,
                Height = 183,
                Margin = new Padding(0, 0, 0, 9),
                BackColor = RedSoft,
                Padding = new Padding(6)
            };
            missingCard.Paint += PaintBorder;
            var missingLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                BackColor = RedSoft,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            missingLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            missingLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            missingTitleLabel = new Label
            {
                Text = "Danh sách khoản mục chưa có giá (0)",
                Dock = DockStyle.Fill,
                ForeColor = Red,
                Font = new Font("Segoe UI", 9.2f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(3, 0, 0, 0)
            };
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
                Font = new Font("Segoe UI", 7.7f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 28,
                GridColor = Color.FromArgb(226, 230, 233),
                DefaultCellStyle =
                {
                    SelectionBackColor = Color.FromArgb(235, 243, 255),
                    SelectionForeColor = TextDark
                },
                ColumnHeadersDefaultCellStyle =
                {
                    BackColor = Color.FromArgb(245, 246, 248),
                    ForeColor = TextDark,
                    Font = new Font("Segoe UI", 7.7f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                RowTemplate = { Height = 27 }
            };
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "STT",
                Width = 38
            });
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Mã hiệu",
                Width = 76
            });
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Tên tài nguyên",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 115
            });
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Đơn vị",
                Width = 52
            });
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Nhóm",
                Width = 74
            });
            missingGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "",
                Width = 24
            });
            missingLayout.Controls.Add(missingTitleLabel, 0, 0);
            missingLayout.Controls.Add(missingGrid, 0, 1);
            missingCard.Controls.Add(missingLayout);
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
                "Từ các bảng khối lượng, định mức",
                EstimateUiIconKind.Clipboard,
                Green,
                GreenSoft);
            Button price = FeatureButton(
                "Cập nhật giá",
                "Lấy giá từ bảng giá hoặc nhập thủ công",
                EstimateUiIconKind.Worker,
                Blue,
                BlueSoft);
            Button labor = FeatureButton(
                "Sinh công thức NC",
                "Tạo công thức tính nhân công từ định mức",
                EstimateUiIconKind.Formula,
                Color.FromArgb(110, 73, 200),
                Color.FromArgb(245, 240, 255));
            Button machine = FeatureButton(
                "Sinh công thức giá ca máy",
                "Tạo công thức tính giá ca máy từ định mức",
                EstimateUiIconKind.Calculator,
                Color.FromArgb(230, 132, 28),
                AmberSoft);

            aggregate.Click += (s, e) => RefreshPreview();
            price.Click += (s, e) => ActivateResourceSheet();
            labor.Click += (s, e) => GenerateResourceSheet();
            machine.Click += (s, e) => GenerateResourceSheet();

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
                Text = "Giá trị trong bảng này được sinh tự động bằng công thức hoặc liên kết từ các bảng giá (VL, nhân công, ca máy), không phải số liệu cố định. Khi cập nhật bảng giá, các giá trị sẽ tự động được cập nhật."
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

                EstimateV2ResourceRequirement[] materials = preview.PriceSheetResources
                    .Where(item => item.Kind == NormResourceKind.Material)
                    .ToArray();
                EstimateV2ResourceRequirement[] labor = preview.PriceSheetResources
                    .Where(item => item.Kind == NormResourceKind.Labor)
                    .ToArray();
                EstimateV2ResourceRequirement[] machines = preview.PriceSheetResources
                    .Where(item => item.Kind == NormResourceKind.Machine)
                    .ToArray();

                materialValue.Text = materials.Length.ToString("N0");
                laborValue.Text = labor.Length.ToString("N0");
                machineValue.Text = machines.Length.ToString("N0");

                var missing = BuildMissingItems();
                missingValue.Text = missing.Count.ToString("N0");
                missingTitleLabel.Text = "Danh sách khoản mục chưa có giá (" +
                    missing.Count.ToString("N0") + ")";
                PopulateMissing(missing);

                UpdateGroupSummary(materialSummary, materials);
                UpdateGroupSummary(laborSummary, labor);
                UpdateGroupSummary(machineSummary, machines);

                if (preview.MissingPackageBindings.Count > 0)
                {
                    ShowStatus(
                        "Có binding chưa resolve được package: " +
                        string.Join(" | ", preview.MissingPackageBindings.Take(2)) +
                        (preview.MissingPackageBindings.Count > 2 ? " ..." : string.Empty),
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
                    string logicalNote = preview.UnresolvedLogicalResources.Count > 0
                        ? " Có " + preview.UnresolvedLogicalResources.Count +
                            " tài nguyên lựa chọn; VL-NC-M đã đưa các phương án giá vào sheet, lựa chọn cụ thể sẽ chốt ở bước đơn giá."
                        : string.Empty;
                    ShowStatus(
                        "Đã gom " + preview.PriceSheetResources.Count +
                        " tài nguyên giá từ " + preview.Plan.BoundWorkItemCount +
                        " công tác và " + preview.Plan.UniqueNormBindingCount +
                        " định mức/variant." + logicalNote, false);
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

            foreach (EstimateV2ResourceRequirement item in preview.PriceSheetResources)
            {
                if (!item.RequiresUnitPrice)
                    continue;

                decimal workbookPrice;
                bool hasPrice = WorkbookEstimateV2ResourceService.TryReadWorkbookUnitPrice(
                    workbook,
                    item,
                    out workbookPrice);
                if (!hasPrice)
                {
                    PriceProfilePrice price;
                    hasPrice = previewProfile != null &&
                        previewProfile.TryFind(item.Code, out price) &&
                        price.AppliedUnitPriceVnd > 0m;
                }
                if (hasPrice)
                    continue;

                result.Add(new MissingItem(
                    item.Code,
                    ResolveName(item.Code),
                    item.Unit,
                    KindText(item.Kind)));
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
            PriceProfilePrice price;
            if (previewProfile != null && previewProfile.TryFind(code, out price))
                return price.Entry.DisplayName;
            return EstimateV2ResourceNames.Get(code);
        }

        private void UpdateGroupSummary(
            GroupSummaryView summary,
            IReadOnlyList<EstimateV2ResourceRequirement> items)
        {
            decimal total = 0m;
            foreach (EstimateV2ResourceRequirement item in items)
            {
                if (!item.RequiresUnitPrice)
                    continue;

                decimal current;
                if (WorkbookEstimateV2ResourceService.TryReadWorkbookUnitPrice(
                    workbook,
                    item,
                    out current))
                {
                    total += current;
                    continue;
                }

                PriceProfilePrice price;
                if (previewProfile != null &&
                    previewProfile.TryFind(item.Code, out price) &&
                    price.AppliedUnitPriceVnd > 0m)
                {
                    total += price.AppliedUnitPriceVnd;
                }
            }

            summary.CountLabel.Text =
                items.Count.ToString("N0") + " khoản mục";
            summary.ValueLabel.Text =
                total.ToString("#,##0", CultureInfo.CurrentCulture);
            summary.UnitLabel.Text = "(đồng)";
        }

        private void PopulateMissing(IEnumerable<MissingItem> items)
        {
            missingGrid.Rows.Clear();
            int index = 1;
            foreach (MissingItem item in items)
            {
                missingGrid.Rows.Add(
                    index,
                    item.Code,
                    item.Name,
                    item.Unit,
                    item.Group,
                    "›");
                index++;
            }
        }

        private void ActivateResourceSheet()
        {
            Excel.Worksheet worksheet = null;
            try
            {
                worksheet =
                    WorkbookEstimateV2CompatibilityService
                        .ResolveOutputWorksheet(
                            workbook,
                            WorksheetRole.ResourcePrices,
                            "VL-NC-M");
                if (worksheet != null)
                {
                    worksheet.Activate();
                    ShowStatus(
                        "Đã mở sheet " +
                        worksheet.Name +
                        ". Writer V2 giữ A:F là vùng in và đặt metadata ngoài vùng in.",
                        false);
                    return;
                }

                ShowStatus(
                    "Workbook chưa có sheet VL-NC-M. Writer sẽ tạo sheet theo đúng mẫu in khi sinh dữ liệu.",
                    true);
            }
            finally
            {
                Release(worksheet);
            }
        }

        private void GenerateResourceSheet()
        {
            try
            {
                WorkbookEstimateV2ResourceWriteResult result =
                    WorkbookEstimateV2ResourceSheetWriter.Apply(workbook);
                RefreshPreview();
                ShowStatus(
                    "Đã cập nhật " + result.WorksheetName +
                    ": " + result.LaborCount + " nhóm NC, " +
                    result.MachineCount + " máy, " +
                    result.MaterialCount + " vật liệu; " +
                    result.FormulaCount + " công thức. " +
                    (result.MissingInputCount > 0
                        ? "Còn " + result.MissingInputCount + " đầu vào cần bổ sung."
                        : "Đủ đầu vào hiện có."),
                    result.MissingInputCount > 0);
                ActivateResourceSheet();
            }
            catch (Exception ex)
            {
                ShowStatus("Không sinh được VL-NC-M: " + ex.Message, true);
            }
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

        private static GroupSummaryView AddGroupCard(
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
                Location = new Point(10, 9),
                Size = new Size(30, 30),
                SizeMode = PictureBoxSizeMode.CenterImage
            });
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(45, 7),
                Height = 20,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = color,
                AutoEllipsis = true
            };
            var countLabel = new Label
            {
                Location = new Point(45, 27),
                Height = 18,
                Font = new Font("Segoe UI", 7.2f),
                ForeColor = TextMuted,
                Text = "0 khoản mục",
                AutoEllipsis = true
            };
            var valueLabel = new Label
            {
                Location = new Point(8, 51),
                Height = 30,
                Font = new Font("Segoe UI", 13.5f, FontStyle.Bold),
                ForeColor = color,
                TextAlign = ContentAlignment.MiddleCenter,
                Text = "0"
            };
            var unitLabel = new Label
            {
                Location = new Point(8, 82),
                Height = 18,
                Font = new Font("Segoe UI", 7.2f),
                ForeColor = TextMuted,
                TextAlign = ContentAlignment.TopCenter,
                Text = "(đồng)"
            };
            card.Controls.Add(titleLabel);
            card.Controls.Add(countLabel);
            card.Controls.Add(valueLabel);
            card.Controls.Add(unitLabel);
            card.Resize += (s, e) =>
            {
                titleLabel.Width = Math.Max(48, card.ClientSize.Width - 52);
                countLabel.Width = Math.Max(48, card.ClientSize.Width - 52);
                valueLabel.Width = Math.Max(60, card.ClientSize.Width - 16);
                unitLabel.Width = Math.Max(60, card.ClientSize.Width - 16);
            };
            parent.Controls.Add(card, column, 0);
            return new GroupSummaryView(countLabel, valueLabel, unitLabel);
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
                case NormResourceKind.Machine: return "Máy thi công";
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

        private sealed class GroupSummaryView
        {
            internal GroupSummaryView(
                Label countLabel,
                Label valueLabel,
                Label unitLabel)
            {
                CountLabel = countLabel;
                ValueLabel = valueLabel;
                UnitLabel = unitLabel;
            }

            internal Label CountLabel { get; }
            internal Label ValueLabel { get; }
            internal Label UnitLabel { get; }
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
