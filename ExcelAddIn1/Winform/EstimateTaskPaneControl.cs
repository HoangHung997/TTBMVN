using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal sealed class EstimateTaskPaneControl : UserControl
    {
        private static readonly Color Green = Color.FromArgb(0, 137, 70);
        private static readonly Color GreenDark = Color.FromArgb(0, 103, 55);
        private static readonly Color GreenSoft = Color.FromArgb(235, 248, 239);
        private static readonly Color Border = Color.FromArgb(222, 228, 223);
        private static readonly Color TextDark = Color.FromArgb(33, 43, 54);
        private static readonly Color TextMuted = Color.FromArgb(92, 103, 112);
        private static readonly Color Amber = Color.FromArgb(244, 166, 35);
        private static readonly Color Red = Color.FromArgb(220, 53, 69);
        private static readonly Color Blue = Color.FromArgb(45, 112, 229);

        private readonly Excel.Workbook workbook;
        private readonly Label projectNameValue;
        private readonly Label projectStatusValue;
        private readonly MetricCard totalCard;
        private readonly MetricCard boundCard;
        private readonly MetricCard sheetCard;
        private readonly MetricCard packageCard;
        private readonly List<StepRow> steps = new List<StepRow>();
        private readonly FlowLayoutPanel sheetTiles;
        private readonly Label footerText;
        private readonly Panel overviewRoot;
        private Control activeChild;

        internal EstimateTaskPaneControl(Excel.Workbook workbook)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;

            overviewRoot = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White
            };
            Controls.Add(overviewRoot);

            TryReconcileOnOpen();

            var content = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                BackColor = Color.White,
                Padding = new Padding(14, 12, 14, 14),
                ColumnCount = 1
            };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            overviewRoot.Controls.Add(content);

            content.Controls.Add(BuildHeader(), 0, content.RowCount++);
            var projectCard = new EstimateCardPanel
            {
                Dock = DockStyle.Top,
                Height = 112,
                Margin = new Padding(0, 10, 0, 8),
                BackColor = GreenSoft,
                BorderColor = Color.FromArgb(205, 232, 211),
                Radius = 11
            };
            var projectIcon = new PictureBox
            {
                Image = EstimateUiIcons.Create(EstimateUiIconKind.Clipboard, 29, Green),
                SizeMode = PictureBoxSizeMode.CenterImage,
                Location = new Point(14, 17),
                Size = new Size(38, 38)
            };
            projectCard.Controls.Add(projectIcon);
            projectNameValue = new Label
            {
                AutoSize = false,
                Location = new Point(62, 13),
                Height = 31,
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
                Font = new Font("Segoe UI", 10.2f, FontStyle.Bold),
                ForeColor = TextDark,
                AutoEllipsis = true
            };
            projectStatusValue = new Label
            {
                AutoSize = false,
                Location = new Point(62, 46),
                Height = 51,
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
                Font = new Font("Segoe UI", 8.8f),
                ForeColor = TextMuted,
                AutoEllipsis = true
            };
            projectCard.Controls.Add(projectNameValue);
            projectCard.Controls.Add(projectStatusValue);
            projectCard.Resize += (s, e) =>
            {
                projectNameValue.Width = Math.Max(120, projectCard.ClientSize.Width - 78);
                projectStatusValue.Width = Math.Max(120, projectCard.ClientSize.Width - 78);
            };
            content.Controls.Add(projectCard, 0, content.RowCount++);

            var metrics = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 104,
                Margin = new Padding(0, 0, 0, 8),
                ColumnCount = 4,
                RowCount = 1
            };
            for (int i = 0; i < 4; i++)
                metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            totalCard = new MetricCard("Công tác", EstimateUiIconKind.Clipboard, Green);
            boundCard = new MetricCard("Đã gắn", EstimateUiIconKind.Check, Green);
            sheetCard = new MetricCard("Sheet in", EstimateUiIconKind.Document, Blue);
            packageCard = new MetricCard("Gói pháp lý", EstimateUiIconKind.Folder, Amber);
            metrics.Controls.Add(totalCard, 0, 0);
            metrics.Controls.Add(boundCard, 1, 0);
            metrics.Controls.Add(sheetCard, 2, 0);
            metrics.Controls.Add(packageCard, 3, 0);
            content.Controls.Add(metrics, 0, content.RowCount++);

            content.Controls.Add(SectionTitle("Quy trình thực hiện dự toán"), 0, content.RowCount++);
            var stepsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Margin = new Padding(0, 4, 0, 8),
                BackColor = Color.White
            };
            AddStep(
                stepsPanel,
                "Công tác",
                "Nhập, tạo, chỉnh sửa danh mục công tác.",
                ShowWorkItems);
            AddStep(
                stepsPanel,
                "Gắn định mức",
                "Gắn định mức cho các công tác.",
                ShowNormBinding);
            AddStep(
                stepsPanel,
                "VL-NC-M",
                "Kiểm tra, cập nhật giá vật liệu, nhân công, máy thi công.",
                ShowResources);
            AddStep(
                stepsPanel,
                "DG Cạn",
                "Tính đơn giá cho điều kiện thi công trên cạn.",
                ShowLandRates);
            AddStep(
                stepsPanel,
                "DG Nước",
                "Tính đơn giá cho điều kiện thi công dưới nước.",
                ShowWaterRates);
            AddStep(
                stepsPanel,
                "THKP-TC & Kiểm tra",
                "Tổng hợp chi phí, cập nhật THKP-TC, kiểm tra dữ liệu.",
                ShowCostSummary);
            stepsPanel.SizeChanged += (s, e) =>
            {
                foreach (Control control in stepsPanel.Controls)
                    control.Width = Math.Max(250, stepsPanel.ClientSize.Width - 4);
            };
            content.Controls.Add(stepsPanel, 0, content.RowCount++);

            content.Controls.Add(SectionTitle("Các bảng tính đầu ra chính"), 0, content.RowCount++);
            sheetTiles = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Margin = new Padding(0, 4, 0, 8),
                BackColor = Color.White
            };
            content.Controls.Add(sheetTiles, 0, content.RowCount++);

            var footer = new EstimateCardPanel
            {
                Dock = DockStyle.Top,
                Height = 58,
                Margin = new Padding(0, 4, 0, 0),
                BackColor = GreenSoft,
                BorderColor = Color.FromArgb(205, 232, 211),
                Radius = 10
            };
            var footerIcon = new PictureBox
            {
                Image = EstimateUiIcons.Create(EstimateUiIconKind.Check, 26, Green),
                Location = new Point(9, 13),
                Size = new Size(31, 31),
                SizeMode = PictureBoxSizeMode.CenterImage
            };
            footerText = new Label
            {
                Location = new Point(46, 8),
                Height = 42,
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
                ForeColor = GreenDark,
                Font = new Font("Segoe UI", 8.6f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            footer.Controls.Add(footerIcon);
            footer.Controls.Add(footerText);
            footer.Resize += (s, e) =>
            {
                footerText.Width = Math.Max(150, footer.ClientSize.Width - 57);
            };
            content.Controls.Add(footer, 0, content.RowCount++);

            RefreshOverview();
        }

        private void TryReconcileOnOpen()
        {
            try
            {
                WorkbookEstimateV2RegistrationService.ReconcileAll(workbook);
            }
            catch (Exception ex)
            {
                // Missing/legacy metadata khong duoc phep chan viec mo task pane.
                RuntimeLogger.Log(ex, "Reconcile Estimate V2 on open");
            }
        }

        internal void RefreshOverview()
        {
            projectNameValue.Text = SafeWorkbookName();

            EstimateV2State state;
            bool hasState = WorkbookEstimateV2StateService.TryLoad(workbook, out state);
            int total = hasState ? state.WorkItems.Count(item => !item.IsOrphaned) : 0;
            int bound = hasState ? state.WorkItems.Count(item => !item.IsOrphaned && item.HasNormBinding) : 0;
            int unbound = Math.Max(0, total - bound);
            int sheets = CountMainOutputSheets();

            ProjectProfile profile;
            bool hasProfile = WorkbookProjectProfileService.TryLoad(workbook, out profile);
            bool packageReady = hasProfile &&
                !string.IsNullOrWhiteSpace(profile.RegulationPackageId);

            projectStatusValue.Text = packageReady
                ? "Workbook đã có hồ sơ dự toán. Gói pháp lý: " +
                    profile.RegulationPackageId + " v" + profile.RegulationPackageVersion + "."
                : "Có thể làm việc ngay. Chưa có gói pháp lý không làm khóa module; chỉ các chức năng tra/gắn định mức cần dữ liệu pháp lý.";

            totalCard.SetValue(total.ToString("N0"), total == 0 ? "chưa đăng ký" : "công tác");
            boundCard.SetValue(bound.ToString("N0"), unbound == 0 && total > 0
                ? "đã gắn đủ"
                : unbound.ToString("N0") + " chưa gắn");
            sheetCard.SetValue(sheets.ToString("N0"), "bảng sẵn sàng");
            packageCard.SetValue(packageReady ? "1" : "0", packageReady ? "đã thiết lập" : "chưa chọn");

            SetStep(0, total > 0 ? StepState.Done : StepState.Ready);
            SetStep(1, total > 0 && unbound == 0 ? StepState.Done :
                total > 0 ? StepState.Warning : StepState.Ready);
            SetStep(2, SheetExists("VL-NC-M") ? StepState.Done : StepState.Ready);
            SetStep(3, SheetExists("DG Can") ? StepState.Done : StepState.Ready);
            SetStep(4, SheetExists("DG Nuoc") || SheetExists("DG Nước") ? StepState.Done : StepState.Ready);
            bool thkpLinked = false;
            try
            {
                thkpLinked =
                    WorkbookEstimateV2CostLinkService
                        .HasDirectCostLinks(workbook);
            }
            catch
            {
            }
            SetStep(
                5,
                thkpLinked
                    ? StepState.Done
                    : StepState.Ready);

            RebuildSheetTiles();
            footerText.Text = "Có thể mở và sử dụng module này ngay cả khi chưa gắn THKP-TC hoặc chưa chọn gói pháp lý.";
        }

        private Control BuildHeader()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 59,
                Margin = new Padding(0),
                BackColor = Color.White
            };
            var title = new Label
            {
                Text = "Tổng quan",
                Location = new Point(0, 0),
                AutoSize = true,
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor = TextDark
            };
            var subtitle = new Label
            {
                Text = "Tổng quan dự án và tiến độ thực hiện dự toán RPBM",
                Location = new Point(1, 33),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.8f),
                ForeColor = TextMuted
            };
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);
            return panel;
        }

        private static Label SectionTitle(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Top,
                Height = 28,
                Margin = new Padding(0, 5, 0, 0),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = TextDark,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private void AddStep(
            FlowLayoutPanel parent,
            string title,
            string detail,
            Action navigateAction)
        {
            var row = new StepRow(steps.Count + 1, title, detail);
            if (navigateAction != null)
                row.NavigateRequested += (s, e) => navigateAction();
            steps.Add(row);
            parent.Controls.Add(row);
        }

        private void ShowWorkItems()
        {
            ShowChild(new EstimateWorkItemsPaneView(
                workbook,
                ShowOverview,
                ShowNormBinding));
        }

        private void ShowNormBinding()
        {
            ShowChild(new EstimateNormBindingPaneView(
                workbook,
                ShowOverview));
        }

        private void ShowResources()
        {
            ShowChild(new EstimateResourcesPaneView(
                workbook,
                ShowOverview));
        }

        private void ShowLandRates()
        {
            ShowChild(new EstimateUnitRatesPaneView(
                workbook,
                EstimateV2RateEnvironment.Land,
                ShowOverview));
        }

        private void ShowWaterRates()
        {
            ShowChild(new EstimateUnitRatesPaneView(
                workbook,
                EstimateV2RateEnvironment.InlandWater,
                ShowOverview,
                ShowSeaRates));
        }

        private void ShowSeaRates()
        {
            ShowChild(new EstimateUnitRatesPaneView(
                workbook,
                EstimateV2RateEnvironment.Sea,
                ShowOverview));
        }

        private void ShowCostSummary()
        {
            ShowChild(new EstimateCostSummaryPaneView(
                workbook,
                ShowOverview));
        }

        private void ShowOverview()
        {
            if (activeChild != null)
            {
                Controls.Remove(activeChild);
                activeChild.Dispose();
                activeChild = null;
            }
            overviewRoot.Visible = true;
            overviewRoot.BringToFront();
            RefreshOverview();
        }

        private void ShowChild(Control child)
        {
            if (child == null)
                throw new ArgumentNullException(nameof(child));
            if (activeChild != null)
            {
                Controls.Remove(activeChild);
                activeChild.Dispose();
            }
            activeChild = child;
            activeChild.Dock = DockStyle.Fill;
            overviewRoot.Visible = false;
            Controls.Add(activeChild);
            activeChild.BringToFront();
        }

        private void SetStep(int index, StepState state)
        {
            if (index >= 0 && index < steps.Count)
                steps[index].SetState(state);
        }

        private void RebuildSheetTiles()
        {
            foreach (Control old in sheetTiles.Controls)
                old.Dispose();
            sheetTiles.Controls.Clear();

            AddSheetTile("THKP-TC", "Tổng hợp chi phí", Color.FromArgb(48, 143, 214));
            AddSheetTile("Gia DT TC", "Bảng dự toán giá", Green);
            AddSheetTile("DG Can", "Đơn giá thi công cạn", Amber);
            AddSheetTile(
                SheetExists("DG Nước") ? "DG Nước" : "DG Nuoc",
                "Đơn giá thi công nước",
                Color.FromArgb(42, 125, 213));
            if (SheetExists("DG Bien") || SheetExists("DG Biển"))
            {
                AddSheetTile(
                    SheetExists("DG Biển") ? "DG Biển" : "DG Bien",
                    "Đơn giá thi công biển",
                    Color.FromArgb(32, 134, 163));
            }
            AddSheetTile("VL-NC-M", "Vật liệu, nhân công, máy", Color.FromArgb(130, 75, 196));
        }

        private void AddSheetTile(string name, string detail, Color color)
        {
            var tile = new EstimateCardPanel
            {
                Width = 122,
                Height = 70,
                Margin = new Padding(0, 0, 7, 7),
                BackColor = Color.White,
                BorderColor = Border,
                Radius = 9
            };
            var icon = new PictureBox
            {
                Image = EstimateUiIcons.Create(EstimateUiIconKind.Document, 22, color),
                Location = new Point(8, 10),
                Size = new Size(27, 27),
                SizeMode = PictureBoxSizeMode.CenterImage
            };
            var title = new Label
            {
                Text = name,
                Location = new Point(39, 8),
                Size = new Size(76, 23),
                Font = new Font("Segoe UI", 8.2f, FontStyle.Bold),
                ForeColor = TextDark,
                AutoEllipsis = true
            };
            var description = new Label
            {
                Text = detail,
                Location = new Point(8, 36),
                Size = new Size(106, 26),
                Font = new Font("Segoe UI", 7.4f),
                ForeColor = TextMuted,
                AutoEllipsis = true
            };
            tile.Controls.Add(icon);
            tile.Controls.Add(title);
            tile.Controls.Add(description);
            sheetTiles.Controls.Add(tile);
        }

        private int CountMainOutputSheets()
        {
            string[] names = { "THKP-TC", "Gia DT TC", "DG Can", "DG Nuoc", "DG Nước", "VL-NC-M", "DG Bien", "DG Biển" };
            var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index] as Excel.Worksheet;
                        if (sheet == null)
                            continue;
                        if (names.Any(name => string.Equals(name, sheet.Name, StringComparison.OrdinalIgnoreCase)))
                            found.Add(sheet.Name);
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
                return found.Count;
            }
            finally
            {
                Release(sheets);
            }
        }

        private bool SheetExists(string name)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index] as Excel.Worksheet;
                        if (sheet != null && string.Equals(
                            sheet.Name,
                            name,
                            StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
                return false;
            }
            finally
            {
                Release(sheets);
            }
        }

        private string SafeWorkbookName()
        {
            try
            {
                return workbook.Name ?? "Hồ sơ dự toán";
            }
            catch (COMException)
            {
                return "Hồ sơ dự toán";
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private enum StepState
        {
            Ready,
            Done,
            Warning
        }

        private sealed class StepRow : EstimateCardPanel
        {
            private readonly Label stateLabel;
            private readonly Label numberLabel;
            private readonly PictureBox stateIcon;
            private readonly Button navigateButton;

            internal event EventHandler NavigateRequested;

            internal StepRow(int number, string title, string detail)
            {
                Height = 57;
                Width = 360;
                Margin = new Padding(0, 0, 0, 6);
                BackColor = Color.White;
                BorderColor = Border;
                Radius = 9;

                var numberCircle = new Panel
                {
                    Location = new Point(8, 10),
                    Size = new Size(35, 35),
                    BackColor = Green
                };
                numberCircle.Paint += (s, e) =>
                {
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var brush = new SolidBrush(Green))
                        e.Graphics.FillEllipse(brush, 0, 0, 34, 34);
                };
                numberLabel = new Label
                {
                    Text = number.ToString(),
                    ForeColor = Color.White,
                    BackColor = Color.Transparent,
                    Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Dock = DockStyle.Fill
                };
                numberCircle.Controls.Add(numberLabel);

                var titleLabel = new Label
                {
                    Text = title,
                    Location = new Point(51, 7),
                    Height = 22,
                    Font = new Font("Segoe UI", 8.8f, FontStyle.Bold),
                    ForeColor = TextDark,
                    AutoEllipsis = true
                };
                var detailLabel = new Label
                {
                    Text = detail,
                    Location = new Point(51, 29),
                    Height = 20,
                    Font = new Font("Segoe UI", 7.5f),
                    ForeColor = TextMuted,
                    AutoEllipsis = true
                };
                stateIcon = new PictureBox
                {
                    Size = new Size(22, 22),
                    SizeMode = PictureBoxSizeMode.CenterImage,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right
                };
                stateLabel = new Label
                {
                    Height = 22,
                    TextAlign = ContentAlignment.MiddleRight,
                    Font = new Font("Segoe UI", 7.4f),
                    Anchor = AnchorStyles.Top | AnchorStyles.Right
                };
                navigateButton = new Button
                {
                    Text = "›",
                    Size = new Size(28, 31),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.White,
                    ForeColor = TextMuted,
                    Font = new Font("Segoe UI", 13f),
                    Cursor = Cursors.Hand,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right
                };
                navigateButton.FlatAppearance.BorderSize = 0;
                navigateButton.Click += (s, e) => NavigateRequested?.Invoke(this, EventArgs.Empty);

                Controls.Add(numberCircle);
                Controls.Add(titleLabel);
                Controls.Add(detailLabel);
                Controls.Add(stateIcon);
                Controls.Add(stateLabel);
                Controls.Add(navigateButton);
                Resize += (s, e) =>
                {
                    titleLabel.Width = Math.Max(120, ClientSize.Width - 150);
                    detailLabel.Width = Math.Max(120, ClientSize.Width - 150);
                    stateIcon.Location = new Point(ClientSize.Width - 120, 17);
                    stateLabel.Location = new Point(ClientSize.Width - 98, 16);
                    stateLabel.Width = 64;
                    navigateButton.Location = new Point(ClientSize.Width - 31, 12);
                };
                SetState(StepState.Ready);
            }

            internal void SetState(StepState state)
            {
                stateIcon.Image?.Dispose();
                switch (state)
                {
                    case StepState.Done:
                        stateIcon.Image = EstimateUiIcons.Create(EstimateUiIconKind.Check, 18, Green);
                        stateLabel.Text = "Đã xong";
                        stateLabel.ForeColor = GreenDark;
                        break;
                    case StepState.Warning:
                        stateIcon.Image = EstimateUiIcons.Create(EstimateUiIconKind.Warning, 18, Amber);
                        stateLabel.Text = "Cần xử lý";
                        stateLabel.ForeColor = Color.DarkGoldenrod;
                        break;
                    default:
                        stateIcon.Image = EstimateUiIcons.Create(EstimateUiIconKind.Refresh, 18, Color.Gray);
                        stateLabel.Text = "Sẵn sàng";
                        stateLabel.ForeColor = TextMuted;
                        break;
                }
            }
        }

        private sealed class MetricCard : EstimateCardPanel
        {
            private readonly Label valueLabel;
            private readonly Label detailLabel;

            internal MetricCard(string title, EstimateUiIconKind iconKind, Color color)
            {
                Margin = new Padding(0, 0, 7, 0);
                BackColor = Color.White;
                BorderColor = Border;
                Radius = 9;
                Dock = DockStyle.Fill;

                var icon = new PictureBox
                {
                    Image = EstimateUiIcons.Create(iconKind, 23, color),
                    Location = new Point(7, 8),
                    Size = new Size(27, 27),
                    SizeMode = PictureBoxSizeMode.CenterImage
                };
                var titleLabel = new Label
                {
                    Text = title,
                    Location = new Point(35, 7),
                    Size = new Size(70, 24),
                    Font = new Font("Segoe UI", 7.2f),
                    ForeColor = TextDark,
                    AutoEllipsis = true
                };
                valueLabel = new Label
                {
                    Location = new Point(8, 35),
                    Size = new Size(96, 30),
                    Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                    ForeColor = color,
                    TextAlign = ContentAlignment.MiddleCenter
                };
                detailLabel = new Label
                {
                    Location = new Point(5, 69),
                    Size = new Size(102, 23),
                    Font = new Font("Segoe UI", 7f),
                    ForeColor = TextMuted,
                    TextAlign = ContentAlignment.TopCenter,
                    AutoEllipsis = true
                };
                Controls.Add(icon);
                Controls.Add(titleLabel);
                Controls.Add(valueLabel);
                Controls.Add(detailLabel);
                Resize += (s, e) =>
                {
                    titleLabel.Width = Math.Max(40, ClientSize.Width - 40);
                    valueLabel.Width = Math.Max(50, ClientSize.Width - 16);
                    detailLabel.Width = Math.Max(50, ClientSize.Width - 10);
                };
            }

            internal void SetValue(string value, string detail)
            {
                valueLabel.Text = value ?? string.Empty;
                detailLabel.Text = detail ?? string.Empty;
            }
        }

        private class EstimateCardPanel : Panel
        {
            internal Color BorderColor { get; set; } = Border;
            internal int Radius { get; set; } = 8;

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle rectangle = new Rectangle(0, 0, Math.Max(1, Width - 1), Math.Max(1, Height - 1));
                using (GraphicsPath path = RoundedRectangle(rectangle, Radius))
                using (var pen = new Pen(BorderColor, 1f))
                    e.Graphics.DrawPath(pen, path);
            }

            private static GraphicsPath RoundedRectangle(Rectangle rect, int radius)
            {
                int diameter = Math.Max(2, radius * 2);
                var path = new GraphicsPath();
                path.AddArc(rect.Left, rect.Top, diameter, diameter, 180, 90);
                path.AddArc(rect.Right - diameter, rect.Top, diameter, diameter, 270, 90);
                path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(rect.Left, rect.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();
                return path;
            }
        }
    }
}
