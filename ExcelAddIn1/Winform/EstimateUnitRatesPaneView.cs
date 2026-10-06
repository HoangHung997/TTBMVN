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
    internal sealed class EstimateUnitRatesPaneView : UserControl
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
        private readonly EstimateV2RateEnvironment environment;
        private readonly Action backAction;
        private readonly Action seaAction;
        private readonly Label neededValue;
        private readonly Label generatedValue;
        private readonly Label missingValue;
        private readonly Label fourthValue;
        private readonly DataGridView rateGrid;
        private readonly DataGridView ratePreviewGrid;
        private readonly Label statusLabel;
        private WorkbookEstimateV2RatePreview preview;

        internal EstimateUnitRatesPaneView(
            Excel.Workbook workbook,
            EstimateV2RateEnvironment environment,
            Action backAction,
            Action seaAction = null)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.environment = environment;
            this.backAction = backAction;
            this.seaAction = seaAction;

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
                Margin = new Padding(0, 4, 0, 9)
            };
            for (int i = 0; i < 4; i++)
                metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            neededValue = AddMetric(
                metrics,
                0,
                environment == EstimateV2RateEnvironment.Land
                    ? "Định mức cần sinh"
                    : environment == EstimateV2RateEnvironment.Sea
                        ? "Định mức biển"
                        : "Định mức nước",
                EstimateUiIconKind.Clipboard,
                Green);
            generatedValue = AddMetric(
                metrics,
                1,
                "Đã sinh",
                EstimateUiIconKind.Check,
                Green);
            missingValue = AddMetric(
                metrics,
                2,
                "Thiếu giá",
                EstimateUiIconKind.Warning,
                Amber);
            fourthValue = AddMetric(
                metrics,
                3,
                environment == EstimateV2RateEnvironment.Land
                    ? "Liên kết công thức"
                    : "Cảnh báo",
                environment == EstimateV2RateEnvironment.Land
                    ? EstimateUiIconKind.Link
                    : EstimateUiIconKind.Error,
                environment == EstimateV2RateEnvironment.Land
                    ? Blue
                    : Red);
            content.Controls.Add(metrics, 0, content.RowCount++);

            if (environment == EstimateV2RateEnvironment.Land)
            {
                BuildLandBody(content, out rateGrid, out ratePreviewGrid);
            }
            else
            {
                BuildWaterBody(content, out rateGrid);
                ratePreviewGrid = null;
            }

            statusLabel = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                MinimumSize = new Size(0, 58),
                Padding = new Padding(10, 9, 10, 9),
                BackColor = BlueSoft,
                ForeColor = Color.FromArgb(35, 88, 180),
                Font = new Font("Segoe UI", 8.1f),
                Text = DefaultInfoText()
            };
            content.Controls.Add(statusLabel, 0, content.RowCount++);

            RefreshPreview();
        }

        internal void RefreshPreview()
        {
            try
            {
                preview = WorkbookEstimateV2RateService.BuildPreview(
                    workbook,
                    environment);

                neededValue.Text = preview.NeededCount.ToString("N0");
                generatedValue.Text = preview.GeneratedCount.ToString("N0");
                missingValue.Text = preview.MissingRateCount.ToString("N0");
                fourthValue.Text = environment == EstimateV2RateEnvironment.Land
                    ? preview.FormulaLinkCount.ToString("N0")
                    : preview.WarningRateCount.ToString("N0");

                PopulateRateGrid();
                if (ratePreviewGrid != null)
                    PopulateRatePreviewGrid();

                if (preview.MissingPackageBindings.Count > 0)
                {
                    ShowStatus(
                        "Có binding chưa resolve được package: " +
                        string.Join(
                            " | ",
                            preview.MissingPackageBindings.Take(2)) +
                        (preview.MissingPackageBindings.Count > 2
                            ? " ..."
                            : string.Empty),
                        true);
                }
                else if (preview.NeededCount == 0)
                {
                    ShowStatus(
                        "Chưa có định mức " +
                        EnvironmentText(environment) +
                        " đang được dùng trong dự toán.",
                        true);
                }
                else
                {
                    ShowStatus(
                        "Đã nhận diện " +
                        preview.NeededCount.ToString("N0") +
                        " đơn giá unique theo định mức + variant; " +
                        preview.GeneratedCount.ToString("N0") +
                        " đã có sheet/link công thức.",
                        false);
                }
            }
            catch (Exception ex)
            {
                neededValue.Text = "0";
                generatedValue.Text = "0";
                missingValue.Text = "!";
                fourthValue.Text = "!";
                rateGrid.Rows.Clear();
                if (ratePreviewGrid != null)
                    ratePreviewGrid.Rows.Clear();
                ShowStatus(
                    "Không đọc được đơn giá: " + ex.Message,
                    true);
            }
        }

        private void BuildLandBody(
            TableLayoutPanel content,
            out DataGridView selectionGrid,
            out DataGridView previewGrid)
        {
            content.Controls.Add(
                StepHeader(
                    "1",
                    "Chọn định mức cần sinh đơn giá",
                    "Chọn các định mức để sinh đơn giá công tác."),
                0,
                content.RowCount++);

            selectionGrid = CreateRateGrid(true);
            var selectionCard = Card(selectionGrid, 166, Color.White);
            content.Controls.Add(selectionCard, 0, content.RowCount++);

            content.Controls.Add(
                StepHeader(
                    "2",
                    "Tùy chọn sinh đơn giá",
                    "Chọn phương thức sinh đơn giá từ định mức và đơn giá VL-NC-M hiện có."),
                0,
                content.RowCount++);

            var options = new Panel
            {
                Dock = DockStyle.Top,
                Height = 86,
                Margin = new Padding(0, 0, 0, 7),
                BackColor = Color.White,
                Padding = new Padding(12, 7, 8, 7)
            };
            options.Paint += PaintBorder;
            AddRadio(
                options,
                "Sinh mới (tạo đơn giá chưa có)",
                5,
                true);
            AddRadio(
                options,
                "Cập nhật từ VL-NC-M (nếu đã có đơn giá)",
                31,
                false);
            AddRadio(
                options,
                "Chỉ sinh định mức đang dùng trong dự toán",
                57,
                false);
            content.Controls.Add(options, 0, content.RowCount++);

            content.Controls.Add(
                StepHeader(
                    "3",
                    "Xem trước đơn giá",
                    "Tóm tắt kết quả dự kiến sau khi sinh đơn giá."),
                0,
                content.RowCount++);

            previewGrid = CreateSummaryGrid();
            content.Controls.Add(
                Card(previewGrid, 128, Color.White),
                0,
                content.RowCount++);

            var generate = new Button
            {
                Text = "⚙  Sinh đơn giá",
                Dock = DockStyle.Top,
                Height = 42,
                Margin = new Padding(0, 0, 0, 9),
                BackColor = Green,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            generate.FlatAppearance.BorderSize = 0;
            generate.Click += (s, e) => GenerateSelectedRates();
            content.Controls.Add(generate, 0, content.RowCount++);
        }

        private void BuildWaterBody(
            TableLayoutPanel content,
            out DataGridView selectionGrid)
        {
            content.Controls.Add(
                SectionTitle("Quy trình thực hiện"),
                0,
                content.RowCount++);

            var process = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Margin = new Padding(0, 2, 0, 8),
                BackColor = Color.White
            };
            AddProcessStep(
                process,
                1,
                environment == EstimateV2RateEnvironment.Sea
                    ? "Chọn công tác trên biển"
                    : "Chọn công tác dưới nước",
                "Chọn hoặc tạo công tác từ định mức chuyên ngành.",
                () => RefreshPreview());
            AddProcessStep(
                process,
                2,
                environment == EstimateV2RateEnvironment.Sea
                    ? "Sinh DG Biển"
                    : "Sinh DG Nước",
                "Tự động lập đơn giá từ định mức, VL-NC-M, máy thi công.",
                GenerateSelectedRates);
            AddProcessStep(
                process,
                3,
                "Cập nhật lại",
                "Cập nhật đơn giá, điều chỉnh hệ số, khối lượng (nếu cần).",
                GenerateSelectedRates);
            AddProcessStep(
                process,
                4,
                "Kiểm tra liên kết",
                "Kiểm tra liên kết giữa định mức, đơn giá và dự toán.",
                CheckLinks);
            process.SizeChanged += (s, e) =>
            {
                foreach (Control control in process.Controls)
                    control.Width = Math.Max(
                        260,
                        process.ClientSize.Width - 2);
            };
            content.Controls.Add(process, 0, content.RowCount++);

            var titleRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = 33,
                Margin = new Padding(0),
                BackColor = Color.White
            };
            var title = new Label
            {
                Text = environment == EstimateV2RateEnvironment.Sea
                    ? "Danh sách định mức công tác trên biển"
                    : "Danh sách định mức công tác dưới nước",
                Location = new Point(0, 4),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.2f, FontStyle.Bold),
                ForeColor = TextDark
            };
            var showAll = new Button
            {
                Text = "Xem tất cả  ›",
                Size = new Size(88, 27),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Blue,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 7.8f),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            showAll.FlatAppearance.BorderSize = 0;
            showAll.Click += (s, e) => RefreshPreview();
            titleRow.Controls.Add(title);
            titleRow.Controls.Add(showAll);
            titleRow.Resize += (s, e) =>
            {
                showAll.Location =
                    new Point(titleRow.ClientSize.Width - 90, 2);
            };
            content.Controls.Add(titleRow, 0, content.RowCount++);

            selectionGrid = CreateRateGrid(false);
            content.Controls.Add(
                Card(selectionGrid, 184, Color.White),
                0,
                content.RowCount++);
        }

        private Control BuildHeader()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 75,
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
                Text = PaneTitle(environment),
                Location = new Point(39, 4),
                AutoSize = true,
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor = TextDark
            };
            var subtitle = new Label
            {
                Text = PaneSubtitle(environment),
                Location = new Point(40, 38),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.4f),
                ForeColor = TextMuted
            };
            panel.Controls.Add(back);
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);

            if (environment == EstimateV2RateEnvironment.InlandWater &&
                seaAction != null)
            {
                var sea = new Button
                {
                    Text = "DG Biển ›",
                    Size = new Size(82, 28),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.White,
                    ForeColor = Blue,
                    Font = new Font("Segoe UI", 7.8f, FontStyle.Bold),
                    Cursor = Cursors.Hand,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right
                };
                sea.FlatAppearance.BorderColor = Border;
                sea.Click += (s, e) => seaAction();
                panel.Controls.Add(sea);
                panel.Resize += (s, e) =>
                {
                    sea.Location =
                        new Point(panel.ClientSize.Width - 86, 3);
                };
            }

            return panel;
        }

        private void PopulateRateGrid()
        {
            rateGrid.Rows.Clear();
            if (preview == null)
                return;

            int index = 1;
            foreach (WorkbookEstimateV2RateItemPreview item in preview.Items)
            {
                string status = item.IsGenerated
                    ? "Đã có"
                    : item.HasMissingPrice
                        ? "Thiếu giá"
                        : "Chưa sinh";

                int rowIndex;
                if (environment == EstimateV2RateEnvironment.Land)
                {
                    rowIndex = rateGrid.Rows.Add(
                        !item.IsGenerated,
                        index,
                        ShortNorm(item.Rate),
                        item.Rate.Title,
                        status);
                }
                else
                {
                    rowIndex = rateGrid.Rows.Add(
                        index,
                        ShortNorm(item.Rate),
                        item.Rate.Title,
                        item.Rate.WorkUnit,
                        status);
                }

                rateGrid.Rows[rowIndex].Tag = item;
                DataGridViewCell statusCell =
                    rateGrid.Rows[rowIndex].Cells[
                        rateGrid.ColumnCount - 1];
                if (item.IsGenerated)
                    statusCell.Style.ForeColor = GreenDark;
                else if (item.HasMissingPrice)
                    statusCell.Style.ForeColor = Color.DarkGoldenrod;
                else
                    statusCell.Style.ForeColor = TextMuted;
                index++;
            }
        }

        private void PopulateRatePreviewGrid()
        {
            ratePreviewGrid.Rows.Clear();
            if (preview == null)
                return;

            int index = 1;
            foreach (WorkbookEstimateV2RateItemPreview item in preview.Items
                .Take(8))
            {
                decimal vl;
                decimal nc;
                decimal machine;
                bool hasVl = WorkbookEstimateV2RateService.TryReadRateComponent(
                    workbook,
                    item.Rate.RateId,
                    "VL",
                    out vl);
                bool hasNc = WorkbookEstimateV2RateService.TryReadRateComponent(
                    workbook,
                    item.Rate.RateId,
                    "NC",
                    out nc);
                bool hasM = WorkbookEstimateV2RateService.TryReadRateComponent(
                    workbook,
                    item.Rate.RateId,
                    "M",
                    out machine);

                decimal total =
                    (hasVl ? vl : 0m) +
                    (hasNc ? nc : 0m) +
                    (hasM ? machine : 0m);

                ratePreviewGrid.Rows.Add(
                    index,
                    ShortNorm(item.Rate),
                    item.Rate.Title,
                    hasVl ? Money(vl) : "-",
                    hasNc ? Money(nc) : "-",
                    hasM ? Money(machine) : "-",
                    item.IsGenerated ? Money(total) : "-");
                index++;
            }
        }

        private void GenerateSelectedRates()
        {
            try
            {
                string[] selected = SelectedRateIds();
                WorkbookEstimateV2RateWriteResult result =
                    WorkbookEstimateV2RateSheetWriter.Apply(
                        workbook,
                        environment,
                        selected);

                RefreshPreview();
                ShowStatus(
                    "Đã cập nhật " +
                    result.WorksheetName +
                    ": " +
                    result.RateCount.ToString("N0") +
                    " đơn giá, " +
                    result.FormulaCount.ToString("N0") +
                    " công thức/link." +
                    (result.MissingPriceRateCount > 0
                        ? " Còn " +
                            result.MissingPriceRateCount.ToString("N0") +
                            " đơn giá có đầu vào giá đang trống."
                        : string.Empty),
                    result.MissingPriceRateCount > 0);
                ActivateGeneratedSheet(result.WorksheetName);
            }
            catch (Exception ex)
            {
                ShowStatus(
                    "Không sinh được đơn giá: " + ex.Message,
                    true);
            }
        }

        private void CheckLinks()
        {
            RefreshPreview();
            if (preview == null)
                return;

            int missing = preview.Items.Count(item => item.HasMissingPrice);
            int warnings = preview.Items.Count(item => item.HasWarning);
            ShowStatus(
                missing == 0 && warnings == 0
                    ? "Liên kết hiện tại không phát hiện thiếu giá/cảnh báo trong preview."
                    : "Kiểm tra: " +
                        missing.ToString("N0") +
                        " đơn giá thiếu đầu vào; " +
                        warnings.ToString("N0") +
                        " đơn giá cần rà soát điều kiện/lựa chọn giá.",
                missing > 0 || warnings > 0);
        }

        private string[] SelectedRateIds()
        {
            if (environment != EstimateV2RateEnvironment.Land)
            {
                return preview?.Items
                    .Select(item => item.Rate.RateId)
                    .ToArray() ?? new string[0];
            }

            var selected = new List<string>();
            foreach (DataGridViewRow row in rateGrid.Rows)
            {
                WorkbookEstimateV2RateItemPreview item =
                    row.Tag as WorkbookEstimateV2RateItemPreview;
                if (item == null)
                    continue;
                object raw = row.Cells[0].Value;
                bool isChecked = raw is bool && (bool)raw;
                if (isChecked)
                    selected.Add(item.Rate.RateId);
            }

            if (selected.Count == 0 && preview != null)
            {
                selected.AddRange(
                    preview.Items
                        .Where(item => !item.IsGenerated)
                        .Select(item => item.Rate.RateId));
            }
            return selected.ToArray();
        }

        private void ActivateGeneratedSheet(string worksheetName)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            try
            {
                sheets = workbook.Worksheets;
                sheet = sheets.Item[worksheetName] as Excel.Worksheet;
                sheet?.Activate();
            }
            catch
            {
            }
            finally
            {
                Release(sheet);
                Release(sheets);
            }
        }

        private DataGridView CreateRateGrid(bool selectable)
        {
            var grid = BaseGrid();
            if (selectable)
            {
                grid.Columns.Add(new DataGridViewCheckBoxColumn
                {
                    HeaderText = "",
                    Width = 30
                });
            }
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "STT",
                Width = 37
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Mã định mức",
                Width = 76
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = selectable
                    ? "Tên định mức"
                    : "Tên công tác",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 120
            });
            if (!selectable)
            {
                grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    HeaderText = "Đơn vị",
                    Width = 54
                });
            }
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Trạng thái",
                Width = 78
            });
            return grid;
        }

        private static DataGridView CreateSummaryGrid()
        {
            var grid = BaseGrid();
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "STT",
                Width = 36
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Mã định mức",
                Width = 76
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Tên định mức",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 110
            });
            foreach (string title in new[]
            {
                "VL (đồng)",
                "NC (đồng)",
                "M (đồng)",
                "Tổng cộng (đồng)"
            })
            {
                grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    HeaderText = title,
                    Width = title.StartsWith("Tổng")
                        ? 90
                        : 68,
                    DefaultCellStyle =
                    {
                        Alignment =
                            DataGridViewContentAlignment.MiddleRight
                    }
                });
            }
            return grid;
        }

        private static DataGridView BaseGrid()
        {
            return new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                Font = new Font("Segoe UI", 7.6f),
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
                    Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                RowTemplate = { Height = 27 }
            };
        }

        private static Panel Card(
            Control child,
            int height,
            Color background)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = height,
                Margin = new Padding(0, 0, 0, 8),
                Padding = new Padding(5),
                BackColor = background
            };
            panel.Paint += PaintBorder;
            child.Dock = DockStyle.Fill;
            panel.Controls.Add(child);
            return panel;
        }

        private static Control StepHeader(
            string number,
            string title,
            string detail)
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 58,
                BackColor = Color.White
            };
            var circle = new Label
            {
                Text = number,
                Location = new Point(0, 8),
                Size = new Size(34, 34),
                BackColor = Green,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(44, 4),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.2f, FontStyle.Bold),
                ForeColor = GreenDark
            };
            var detailLabel = new Label
            {
                Text = detail,
                Location = new Point(44, 27),
                Height = 28,
                Width = 320,
                Font = new Font("Segoe UI", 7.8f),
                ForeColor = TextMuted,
                AutoEllipsis = true
            };
            panel.Controls.Add(circle);
            panel.Controls.Add(titleLabel);
            panel.Controls.Add(detailLabel);
            return panel;
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
                Location = new Point(8, 10),
                Size = new Size(28, 28),
                SizeMode = PictureBoxSizeMode.CenterImage
            });
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(38, 8),
                Height = 26,
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

        private static void AddRadio(
            Control parent,
            string text,
            int top,
            bool isChecked)
        {
            parent.Controls.Add(new RadioButton
            {
                Text = text,
                Location = new Point(12, top),
                Width = 330,
                Height = 24,
                Checked = isChecked,
                Font = new Font("Segoe UI", 8.1f),
                ForeColor = TextDark
            });
        }

        private static void AddProcessStep(
            FlowLayoutPanel parent,
            int number,
            string title,
            string detail,
            Action action)
        {
            var row = new Panel
            {
                Height = 54,
                Width = 360,
                Margin = new Padding(0, 0, 0, 4),
                BackColor = Color.White
            };
            row.Paint += PaintBorder;

            var circle = new Label
            {
                Text = number.ToString(
                    CultureInfo.InvariantCulture),
                Location = new Point(9, 9),
                Size = new Size(34, 34),
                BackColor = Green,
                ForeColor = Color.White,
                Font = new Font(
                    "Segoe UI",
                    10f,
                    FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(51, 6),
                Height = 21,
                Font = new Font(
                    "Segoe UI",
                    8.6f,
                    FontStyle.Bold),
                ForeColor = TextDark,
                AutoEllipsis = true
            };
            var detailLabel = new Label
            {
                Text = detail,
                Location = new Point(51, 27),
                Height = 20,
                Font = new Font("Segoe UI", 7.4f),
                ForeColor = TextMuted,
                AutoEllipsis = true
            };
            var button = new Button
            {
                Text = "Thực hiện  ›",
                Size = new Size(88, 31),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = TextDark,
                Font = new Font("Segoe UI", 7.6f),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            button.FlatAppearance.BorderColor = Border;
            button.Click += (s, e) => action?.Invoke();

            row.Controls.Add(circle);
            row.Controls.Add(titleLabel);
            row.Controls.Add(detailLabel);
            row.Controls.Add(button);
            row.Resize += (s, e) =>
            {
                titleLabel.Width =
                    Math.Max(110, row.ClientSize.Width - 160);
                detailLabel.Width =
                    Math.Max(110, row.ClientSize.Width - 160);
                button.Location =
                    new Point(row.ClientSize.Width - 97, 11);
            };
            parent.Controls.Add(row);
        }

        private void ShowStatus(
            string message,
            bool warning)
        {
            statusLabel.BackColor =
                warning ? AmberSoft : BlueSoft;
            statusLabel.ForeColor =
                warning
                    ? Color.DarkGoldenrod
                    : Color.FromArgb(35, 88, 180);
            statusLabel.Text = message;
        }

        private string DefaultInfoText()
        {
            if (environment == EstimateV2RateEnvironment.Land)
            {
                return "Mỗi định mức chỉ được sinh một block đơn giá và được sử dụng chung cho tất cả công tác liên kết đến định mức đó.";
            }
            if (environment == EstimateV2RateEnvironment.Sea)
            {
                return "Sheet DG Biển chỉ được tạo hoặc cập nhật khi dự toán có sử dụng định mức công tác trên biển.";
            }
            return "Sheet DG Nước chỉ được tạo hoặc cập nhật khi dự toán có sử dụng các định mức công tác dưới nước.";
        }

        private static string PaneTitle(
            EstimateV2RateEnvironment environment)
        {
            switch (environment)
            {
                case EstimateV2RateEnvironment.Land:
                    return "DG Cạn";
                case EstimateV2RateEnvironment.InlandWater:
                    return "DG Nước";
                case EstimateV2RateEnvironment.Sea:
                    return "DG Biển";
                default:
                    return "Đơn giá";
            }
        }

        private static string PaneSubtitle(
            EstimateV2RateEnvironment environment)
        {
            switch (environment)
            {
                case EstimateV2RateEnvironment.Land:
                    return "Hỗ trợ sinh và quản lý đơn giá công tác trên cạn";
                case EstimateV2RateEnvironment.InlandWater:
                    return "Hỗ trợ lập đơn giá công tác dưới nước";
                case EstimateV2RateEnvironment.Sea:
                    return "Hỗ trợ lập đơn giá công tác trên biển";
                default:
                    return string.Empty;
            }
        }

        private static string EnvironmentText(
            EstimateV2RateEnvironment environment)
        {
            switch (environment)
            {
                case EstimateV2RateEnvironment.Land:
                    return "trên cạn";
                case EstimateV2RateEnvironment.InlandWater:
                    return "dưới nước";
                case EstimateV2RateEnvironment.Sea:
                    return "trên biển";
                default:
                    return environment.ToString();
            }
        }

        private static string ShortNorm(
            EstimateV2RateItem rate)
        {
            string code = rate.NormCode ?? string.Empty;
            if (code.StartsWith(
                "NORM-",
                StringComparison.OrdinalIgnoreCase))
            {
                code = code.Substring(5);
            }
            return code + " / " + rate.VariantCode;
        }

        private static string Money(decimal value)
        {
            return value.ToString(
                "#,##0",
                CultureInfo.CurrentCulture);
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

        private static void Release(object value)
        {
            if (value != null &&
                System.Runtime.InteropServices.Marshal.IsComObject(value))
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
            }
        }
    }
}
