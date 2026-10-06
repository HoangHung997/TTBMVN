using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal sealed class EstimateSettingsPaneView : UserControl
    {
        private static readonly Color Green =
            Color.FromArgb(0, 137, 70);
        private static readonly Color GreenDark =
            Color.FromArgb(0, 103, 55);
        private static readonly Color GreenSoft =
            Color.FromArgb(235, 248, 239);
        private static readonly Color Blue =
            Color.FromArgb(45, 112, 229);
        private static readonly Color Amber =
            Color.FromArgb(244, 166, 35);
        private static readonly Color Border =
            Color.FromArgb(222, 228, 223);
        private static readonly Color TextDark =
            Color.FromArgb(33, 43, 54);
        private static readonly Color TextMuted =
            Color.FromArgb(92, 103, 112);

        private readonly Excel.Workbook workbook;
        private readonly Action backAction;
        private readonly Action<EstimateV2Settings>
            settingsChangedAction;

        private readonly Label workbookValue;
        private readonly Label workbookDetail;
        private readonly Label roleValue;
        private readonly Label roleDetail;
        private readonly Label metadataValue;
        private readonly Label metadataDetail;
        private readonly Label autosaveValue;
        private readonly Label autosaveDetail;

        private readonly Dictionary<
            EstimateV2OutputSlot,
            ComboBox> mappingCombos =
                new Dictionary<
                    EstimateV2OutputSlot,
                    ComboBox>();

        private readonly ToggleSwitch autoRestoreToggle;
        private readonly ToggleSwitch validateOnOpenToggle;
        private readonly ToggleSwitch formulaLinkToggle;
        private readonly ToggleSwitch autoSyncToggle;
        private readonly ToggleSwitch hiddenColumnsToggle;
        private readonly ToggleSwitch warnMappingToggle;
        private readonly Label statusLabel;

        private WorkbookEstimateV2SettingsSnapshot snapshot;

        internal EstimateSettingsPaneView(
            Excel.Workbook workbook,
            Action backAction,
            Action<EstimateV2Settings>
                settingsChangedAction)
        {
            this.workbook = workbook ??
                throw new ArgumentNullException(nameof(workbook));
            this.backAction = backAction;
            this.settingsChangedAction =
                settingsChangedAction;

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
            content.ColumnStyles.Add(
                new ColumnStyle(
                    SizeType.Percent,
                    100));
            root.Controls.Add(content);

            content.Controls.Add(
                BuildHeader(),
                0,
                content.RowCount++);

            var metrics = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 93,
                ColumnCount = 4,
                Margin = new Padding(0, 4, 0, 9)
            };
            for (int index = 0;
                index < 4;
                index++)
            {
                metrics.ColumnStyles.Add(
                    new ColumnStyle(
                        SizeType.Percent,
                        25));
            }

            MetricLabels workbookMetric =
                AddMetric(
                    metrics,
                    0,
                    "Workbook",
                    EstimateUiIconKind.Document,
                    Green);
            workbookValue =
                workbookMetric.Value;
            workbookDetail =
                workbookMetric.Detail;

            MetricLabels roleMetric =
                AddMetric(
                    metrics,
                    1,
                    "Vai trò sheet",
                    EstimateUiIconKind.Materials,
                    Blue);
            roleValue =
                roleMetric.Value;
            roleDetail =
                roleMetric.Detail;

            MetricLabels metadataMetric =
                AddMetric(
                    metrics,
                    2,
                    "Metadata",
                    EstimateUiIconKind.Document,
                    Blue);
            metadataValue =
                metadataMetric.Value;
            metadataDetail =
                metadataMetric.Detail;

            MetricLabels autosaveMetric =
                AddMetric(
                    metrics,
                    3,
                    "Tự động lưu",
                    EstimateUiIconKind.Cloud,
                    Blue);
            autosaveValue =
                autosaveMetric.Value;
            autosaveDetail =
                autosaveMetric.Detail;

            content.Controls.Add(
                metrics,
                0,
                content.RowCount++);

            content.Controls.Add(
                SectionTitle(
                    "1",
                    "Sheet đầu ra",
                    "Chọn và cấu hình các sheet được Trợ lý Dự toán sử dụng"),
                0,
                content.RowCount++);

            var mappingCard =
                new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 150,
                    BackColor = Color.White,
                    Margin = new Padding(
                        0, 1, 0, 8),
                    Padding = new Padding(7)
                };
            mappingCard.Paint +=
                PaintBorder;

            var mappingGrid =
                new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 2,
                    RowCount = 3,
                    BackColor = Color.White
                };
            mappingGrid.ColumnStyles.Add(
                new ColumnStyle(
                    SizeType.Percent,
                    50));
            mappingGrid.ColumnStyles.Add(
                new ColumnStyle(
                    SizeType.Percent,
                    50));
            for (int row = 0;
                row < 3;
                row++)
            {
                mappingGrid.RowStyles.Add(
                    new RowStyle(
                        SizeType.Percent,
                        33.333f));
            }

            AddMappingCell(
                mappingGrid,
                0,
                0,
                EstimateV2OutputSlot.CostSummary,
                "THKP-TC",
                "Tổng hợp chi phí");
            AddMappingCell(
                mappingGrid,
                0,
                1,
                EstimateV2OutputSlot.UnitRateWater,
                "DG Nước",
                "Đơn giá nước");
            AddMappingCell(
                mappingGrid,
                1,
                0,
                EstimateV2OutputSlot.EstimateAppendix,
                "Gia DT TC",
                "Bảng dự toán");
            AddMappingCell(
                mappingGrid,
                1,
                1,
                EstimateV2OutputSlot.ResourcePrices,
                "VL-NC-M",
                "Vật liệu - Nhân công - Máy");
            AddMappingCell(
                mappingGrid,
                2,
                0,
                EstimateV2OutputSlot.UnitRateLand,
                "DG Cạn",
                "Đơn giá cạn");
            AddMappingCell(
                mappingGrid,
                2,
                1,
                EstimateV2OutputSlot.UnitRateSea,
                "DG Biển",
                "Đơn giá biển");

            mappingCard.Controls.Add(
                mappingGrid);
            content.Controls.Add(
                mappingCard,
                0,
                content.RowCount++);

            content.Controls.Add(
                SectionTitle(
                    "2",
                    "Hành vi cập nhật",
                    "Cấu hình cách Trợ lý Dự toán xử lý và cập nhật dữ liệu"),
                0,
                content.RowCount++);

            var behaviorCard =
                new FlowLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    FlowDirection =
                        FlowDirection.TopDown,
                    WrapContents = false,
                    BackColor = Color.White,
                    Margin = new Padding(
                        0, 1, 0, 8),
                    Padding = new Padding(0)
                };
            behaviorCard.SizeChanged +=
                (s, e) =>
                {
                    foreach (Control control in
                        behaviorCard.Controls)
                    {
                        control.Width =
                            Math.Max(
                                260,
                                behaviorCard
                                    .ClientSize.Width -
                                2);
                    }
                };

            autoRestoreToggle =
                AddToggleRow(
                    behaviorCard,
                    "Tự phục hồi ô định mức hiển thị",
                    "Tự động khôi phục công thức/giá trị hiển thị khi bị xóa nhầm",
                    false);
            validateOnOpenToggle =
                AddToggleRow(
                    behaviorCard,
                    "Kiểm tra khi mở file",
                    "Tự động kiểm tra cấu trúc, định dạng và liên kết dữ liệu",
                    false);
            formulaLinkToggle =
                AddToggleRow(
                    behaviorCard,
                    "Tạo công thức thay vì số chết",
                    "Ưu tiên tạo công thức để dễ dàng cập nhật khi có thay đổi",
                    true);
            autoSyncToggle =
                AddToggleRow(
                    behaviorCard,
                    "Tự động đồng bộ khi chèn/xóa dòng",
                    "Giữ liên kết công thức, định mức và metadata",
                    false);

            content.Controls.Add(
                behaviorCard,
                0,
                content.RowCount++);

            content.Controls.Add(
                SectionTitle(
                    "3",
                    "Bảo vệ dữ liệu",
                    "Thiết lập các tùy chọn bảo vệ và kiểm soát dữ liệu quan trọng"),
                0,
                content.RowCount++);

            var protectionCard =
                new FlowLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    FlowDirection =
                        FlowDirection.TopDown,
                    WrapContents = false,
                    BackColor = Color.White,
                    Margin = new Padding(
                        0, 1, 0, 8),
                    Padding = new Padding(0)
                };
            protectionCard.SizeChanged +=
                (s, e) =>
                {
                    foreach (Control control in
                        protectionCard.Controls)
                    {
                        control.Width =
                            Math.Max(
                                260,
                                protectionCard
                                    .ClientSize.Width -
                                2);
                    }
                };

            protectionCard.Controls.Add(
                BuildFixedProtectionRow(
                    "Sử dụng vùng Custom XML",
                    "Lưu thông tin cấu hình, mapping trong Custom XML ẩn",
                    EstimateUiIconKind.Lock));

            hiddenColumnsToggle =
                AddToggleRow(
                    protectionCard,
                    "Cột kỹ thuật ẩn",
                    "Ẩn các cột mã, định danh, mapping khỏi người dùng",
                    true);

            warnMappingToggle =
                AddToggleRow(
                    protectionCard,
                    "Cảnh báo khi phát hiện xóa mapping",
                    "Hiển thị cảnh báo nếu có thao tác xóa/mất liên kết định mức",
                    false);

            content.Controls.Add(
                protectionCard,
                0,
                content.RowCount++);

            var buttons =
                new TableLayoutPanel
                {
                    Dock = DockStyle.Top,
                    Height = 44,
                    ColumnCount = 3,
                    Margin = new Padding(
                        0, 1, 0, 8)
                };
            buttons.ColumnStyles.Add(
                new ColumnStyle(
                    SizeType.Percent,
                    32));
            buttons.ColumnStyles.Add(
                new ColumnStyle(
                    SizeType.Percent,
                    34));
            buttons.ColumnStyles.Add(
                new ColumnStyle(
                    SizeType.Percent,
                    34));

            Button save =
                ActionButton(
                    "Lưu thiết lập",
                    EstimateUiIconKind.Save,
                    Green,
                    Color.White,
                    true);
            Button reset =
                ActionButton(
                    "Khôi phục mặc định",
                    EstimateUiIconKind.Refresh,
                    Color.White,
                    TextDark,
                    false);
            Button openFolder =
                ActionButton(
                    "Mở thư mục cấu hình",
                    EstimateUiIconKind.Folder,
                    Color.White,
                    TextDark,
                    false);

            save.Click +=
                (s, e) => SaveSettings();
            reset.Click +=
                (s, e) => ResetSettings();
            openFolder.Click +=
                (s, e) => OpenConfigurationFolder();

            buttons.Controls.Add(
                save,
                0,
                0);
            buttons.Controls.Add(
                reset,
                1,
                0);
            buttons.Controls.Add(
                openFolder,
                2,
                0);

            content.Controls.Add(
                buttons,
                0,
                content.RowCount++);

            statusLabel =
                new Label
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    MinimumSize =
                        new Size(0, 50),
                    Padding =
                        new Padding(
                            10, 8, 10, 8),
                    BackColor = GreenSoft,
                    ForeColor = GreenDark,
                    Font =
                        new Font(
                            "Segoe UI",
                            7.9f),
                    Text =
                        "Thiết lập được lưu trong workbook; không tạo thêm sheet kỹ thuật."
                };
            content.Controls.Add(
                statusLabel,
                0,
                content.RowCount++);

            LoadSnapshot();
        }

        private Control BuildHeader()
        {
            var panel =
                new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 72,
                    BackColor = Color.White
                };

            var icon =
                new PictureBox
                {
                    Image =
                        EstimateUiIcons.Create(
                            EstimateUiIconKind.Settings,
                            30,
                            Color.FromArgb(
                                28, 56, 82)),
                    Location = new Point(0, 14),
                    Size = new Size(38, 38),
                    SizeMode =
                        PictureBoxSizeMode.CenterImage
                };

            var overline =
                new Label
                {
                    Text = "Trợ lý Dự toán",
                    Location = new Point(43, 2),
                    AutoSize = true,
                    Font =
                        new Font(
                            "Segoe UI",
                            9f,
                            FontStyle.Bold),
                    ForeColor =
                        Color.FromArgb(
                            28, 56, 82)
                };

            var title =
                new Label
                {
                    Text = "Thiết lập chung",
                    Location = new Point(43, 23),
                    AutoSize = true,
                    Font =
                        new Font(
                            "Segoe UI",
                            14.2f,
                            FontStyle.Bold),
                    ForeColor = TextDark
                };

            var subtitle =
                new Label
                {
                    Text =
                        "Cấu hình workbook và hành vi của Trợ lý Dự toán",
                    Location = new Point(44, 51),
                    Height = 20,
                    Font =
                        new Font(
                            "Segoe UI",
                            8.1f),
                    ForeColor = TextMuted,
                    AutoEllipsis = true
                };

            var back =
                new Button
                {
                    Text = "‹",
                    Size = new Size(29, 29),
                    FlatStyle =
                        FlatStyle.Flat,
                    BackColor =
                        Color.White,
                    ForeColor =
                        TextMuted,
                    Font =
                        new Font(
                            "Segoe UI",
                            14f),
                    Cursor =
                        Cursors.Hand,
                    Anchor =
                        AnchorStyles.Top |
                        AnchorStyles.Right
                };
            back.FlatAppearance.BorderSize = 0;
            back.Click +=
                (s, e) => backAction?.Invoke();

            panel.Controls.Add(icon);
            panel.Controls.Add(overline);
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);
            panel.Controls.Add(back);
            panel.Resize +=
                (s, e) =>
                {
                    subtitle.Width =
                        Math.Max(
                            150,
                            panel.ClientSize.Width -
                            84);
                    back.Location =
                        new Point(
                            panel.ClientSize.Width -
                            31,
                            3);
                };

            return panel;
        }

        private static MetricLabels AddMetric(
            TableLayoutPanel parent,
            int column,
            string title,
            EstimateUiIconKind iconKind,
            Color color)
        {
            var card =
                new Panel
                {
                    Dock = DockStyle.Fill,
                    Margin =
                        new Padding(
                            0, 0, 6, 0),
                    BackColor =
                        Color.White
                };
            card.Paint += PaintBorder;

            var icon =
                new PictureBox
                {
                    Image =
                        EstimateUiIcons.Create(
                            iconKind,
                            22,
                            color),
                    Location = new Point(7, 7),
                    Size = new Size(26, 26),
                    SizeMode =
                        PictureBoxSizeMode.CenterImage
                };
            var titleLabel =
                new Label
                {
                    Text = title,
                    Location = new Point(34, 5),
                    Height = 23,
                    Font =
                        new Font(
                            "Segoe UI",
                            6.9f),
                    ForeColor = TextDark,
                    AutoEllipsis = true
                };
            var value =
                new Label
                {
                    Location = new Point(6, 34),
                    Height = 25,
                    Font =
                        new Font(
                            "Segoe UI",
                            10.5f,
                            FontStyle.Bold),
                    ForeColor = color,
                    TextAlign =
                        ContentAlignment.MiddleCenter,
                    AutoEllipsis = true
                };
            var detail =
                new Label
                {
                    Location = new Point(6, 61),
                    Height = 20,
                    Font =
                        new Font(
                            "Segoe UI",
                            6.7f),
                    ForeColor = TextMuted,
                    TextAlign =
                        ContentAlignment.TopCenter,
                    AutoEllipsis = true
                };

            card.Controls.Add(icon);
            card.Controls.Add(titleLabel);
            card.Controls.Add(value);
            card.Controls.Add(detail);
            card.Resize +=
                (s, e) =>
                {
                    titleLabel.Width =
                        Math.Max(
                            34,
                            card.ClientSize.Width -
                            38);
                    value.Width =
                        Math.Max(
                            50,
                            card.ClientSize.Width -
                            12);
                    detail.Width =
                        Math.Max(
                            50,
                            card.ClientSize.Width -
                            12);
                };

            parent.Controls.Add(
                card,
                column,
                0);
            return new MetricLabels(
                value,
                detail);
        }

        private static Control SectionTitle(
            string number,
            string title,
            string subtitle)
        {
            var panel =
                new Panel
                {
                    Dock = DockStyle.Top,
                    Height = 58,
                    BackColor = Color.White
                };

            var circle =
                new Label
                {
                    Text = number,
                    Location = new Point(0, 9),
                    Size = new Size(34, 34),
                    BackColor = Green,
                    ForeColor = Color.White,
                    Font =
                        new Font(
                            "Segoe UI",
                            10f,
                            FontStyle.Bold),
                    TextAlign =
                        ContentAlignment.MiddleCenter
                };

            var titleLabel =
                new Label
                {
                    Text = title,
                    Location = new Point(43, 5),
                    AutoSize = true,
                    Font =
                        new Font(
                            "Segoe UI",
                            9.2f,
                            FontStyle.Bold),
                    ForeColor = TextDark
                };

            var subtitleLabel =
                new Label
                {
                    Text = subtitle,
                    Location = new Point(43, 28),
                    Height = 24,
                    Width = 330,
                    Font =
                        new Font(
                            "Segoe UI",
                            7.4f),
                    ForeColor = TextMuted,
                    AutoEllipsis = true
                };

            panel.Controls.Add(circle);
            panel.Controls.Add(titleLabel);
            panel.Controls.Add(subtitleLabel);
            panel.Resize +=
                (s, e) =>
                {
                    subtitleLabel.Width =
                        Math.Max(
                            150,
                            panel.ClientSize.Width -
                            48);
                };
            return panel;
        }

        private void AddMappingCell(
            TableLayoutPanel grid,
            int row,
            int column,
            EstimateV2OutputSlot slot,
            string title,
            string detail)
        {
            var panel =
                new Panel
                {
                    Dock = DockStyle.Fill,
                    Margin =
                        new Padding(
                            column == 0 ? 0 : 4,
                            0,
                            column == 0 ? 4 : 0,
                            4),
                    BackColor = Color.White
                };

            var check =
                new PictureBox
                {
                    Image =
                        EstimateUiIcons.Create(
                            EstimateUiIconKind.Check,
                            18,
                            Green),
                    Location = new Point(2, 7),
                    Size = new Size(21, 21),
                    SizeMode =
                        PictureBoxSizeMode.CenterImage
                };

            var label =
                new Label
                {
                    Text = title,
                    Location = new Point(27, 3),
                    Height = 22,
                    Width = 83,
                    Font =
                        new Font(
                            "Segoe UI",
                            7.4f,
                            FontStyle.Bold),
                    ForeColor = TextDark,
                    AutoEllipsis = true
                };

            var combo =
                new ComboBox
                {
                    DropDownStyle =
                        ComboBoxStyle.DropDownList,
                    FlatStyle =
                        FlatStyle.Flat,
                    Location =
                        new Point(105, 1),
                    Height = 26,
                    Font =
                        new Font(
                            "Segoe UI",
                            7f),
                    Anchor =
                        AnchorStyles.Top |
                        AnchorStyles.Left |
                        AnchorStyles.Right
                };

            var description =
                new Label
                {
                    Text = detail,
                    Location = new Point(27, 27),
                    Height = 17,
                    Font =
                        new Font(
                            "Segoe UI",
                            6.5f),
                    ForeColor = TextMuted,
                    AutoEllipsis = true
                };

            panel.Controls.Add(check);
            panel.Controls.Add(label);
            panel.Controls.Add(combo);
            panel.Controls.Add(description);
            panel.Resize +=
                (s, e) =>
                {
                    combo.Width =
                        Math.Max(
                            70,
                            panel.ClientSize.Width -
                            109);
                    description.Width =
                        Math.Max(
                            80,
                            panel.ClientSize.Width -
                            30);
                };

            mappingCombos[slot] = combo;
            grid.Controls.Add(
                panel,
                column,
                row);
        }

        private static ToggleSwitch AddToggleRow(
            FlowLayoutPanel parent,
            string title,
            string detail,
            bool locked)
        {
            var row =
                new Panel
                {
                    Width = 380,
                    Height = 52,
                    Margin =
                        new Padding(
                            0, 0, 0, 3),
                    BackColor = Color.White
                };
            row.Paint += PaintBorder;

            var info =
                new PictureBox
                {
                    Image =
                        EstimateUiIcons.Create(
                            EstimateUiIconKind.Settings,
                            17,
                            Color.FromArgb(
                                49, 78, 104)),
                    Location = new Point(8, 15),
                    Size = new Size(22, 22),
                    SizeMode =
                        PictureBoxSizeMode.CenterImage
                };

            var titleLabel =
                new Label
                {
                    Text = title,
                    Location = new Point(36, 5),
                    Height = 21,
                    Font =
                        new Font(
                            "Segoe UI",
                            7.5f,
                            FontStyle.Bold),
                    ForeColor = TextDark,
                    AutoEllipsis = true
                };

            var detailLabel =
                new Label
                {
                    Text = detail,
                    Location = new Point(36, 27),
                    Height = 19,
                    Font =
                        new Font(
                            "Segoe UI",
                            6.7f),
                    ForeColor = TextMuted,
                    AutoEllipsis = true
                };

            var toggle =
                new ToggleSwitch
                {
                    Width = 38,
                    Height = 22,
                    Anchor =
                        AnchorStyles.Top |
                        AnchorStyles.Right,
                    Locked = locked
                };

            row.Controls.Add(info);
            row.Controls.Add(titleLabel);
            row.Controls.Add(detailLabel);
            row.Controls.Add(toggle);
            row.Resize +=
                (s, e) =>
                {
                    toggle.Location =
                        new Point(
                            row.ClientSize.Width -
                            46,
                            15);
                    titleLabel.Width =
                        Math.Max(
                            100,
                            row.ClientSize.Width -
                            92);
                    detailLabel.Width =
                        Math.Max(
                            100,
                            row.ClientSize.Width -
                            92);
                };

            parent.Controls.Add(row);
            return toggle;
        }

        private static Control BuildFixedProtectionRow(
            string title,
            string detail,
            EstimateUiIconKind iconKind)
        {
            var row =
                new Panel
                {
                    Width = 380,
                    Height = 52,
                    Margin =
                        new Padding(
                            0, 0, 0, 3),
                    BackColor = Color.White
                };
            row.Paint += PaintBorder;

            row.Controls.Add(
                new PictureBox
                {
                    Image =
                        EstimateUiIcons.Create(
                            iconKind,
                            21,
                            Color.FromArgb(
                                49, 78, 104)),
                    Location = new Point(8, 14),
                    Size = new Size(24, 24),
                    SizeMode =
                        PictureBoxSizeMode.CenterImage
                });

            var titleLabel =
                new Label
                {
                    Text = title,
                    Location = new Point(39, 5),
                    Height = 21,
                    Font =
                        new Font(
                            "Segoe UI",
                            7.5f,
                            FontStyle.Bold),
                    ForeColor = TextDark,
                    AutoEllipsis = true
                };
            var detailLabel =
                new Label
                {
                    Text = detail,
                    Location = new Point(39, 27),
                    Height = 19,
                    Font =
                        new Font(
                            "Segoe UI",
                            6.7f),
                    ForeColor = TextMuted,
                    AutoEllipsis = true
                };
            var lockLabel =
                new PictureBox
                {
                    Image =
                        EstimateUiIcons.Create(
                            EstimateUiIconKind.Lock,
                            19,
                            Color.FromArgb(
                                49, 78, 104)),
                    Width = 34,
                    Height = 30,
                    SizeMode =
                        PictureBoxSizeMode.CenterImage,
                    Anchor =
                        AnchorStyles.Top |
                        AnchorStyles.Right
                };

            row.Controls.Add(titleLabel);
            row.Controls.Add(detailLabel);
            row.Controls.Add(lockLabel);
            row.Resize +=
                (s, e) =>
                {
                    lockLabel.Location =
                        new Point(
                            row.ClientSize.Width -
                            42,
                            11);
                    titleLabel.Width =
                        Math.Max(
                            120,
                            row.ClientSize.Width -
                            91);
                    detailLabel.Width =
                        Math.Max(
                            120,
                            row.ClientSize.Width -
                            91);
                };
            return row;
        }

        private static Button ActionButton(
            string text,
            EstimateUiIconKind iconKind,
            Color background,
            Color foreground,
            bool bold)
        {
            var button =
                new Button
                {
                    Text = text,
                    Dock = DockStyle.Fill,
                    Margin =
                        new Padding(
                            0, 0, 6, 0),
                    FlatStyle =
                        FlatStyle.Flat,
                    BackColor = background,
                    ForeColor = foreground,
                    Font =
                        new Font(
                            "Segoe UI",
                            7.2f,
                            bold
                                ? FontStyle.Bold
                                : FontStyle.Regular),
                    Cursor = Cursors.Hand,
                    Image =
                        EstimateUiIcons.Create(
                            iconKind,
                            17,
                            background == Green
                                ? Color.White
                                : GreenDark),
                    ImageAlign =
                        ContentAlignment.MiddleLeft,
                    TextImageRelation =
                        TextImageRelation.ImageBeforeText
                };
            button.FlatAppearance.BorderColor =
                Border;
            button.FlatAppearance.BorderSize =
                background == Green
                    ? 0
                    : 1;
            return button;
        }

        private void LoadSnapshot()
        {
            try
            {
                snapshot =
                    WorkbookEstimateV2SettingsService
                        .CaptureSnapshot(workbook);
                ApplySnapshot(snapshot);
                ShowStatus(
                    "Thiết lập được lưu trong workbook; không tạo thêm sheet kỹ thuật.",
                    false);
            }
            catch (Exception ex)
            {
                snapshot = null;
                ShowStatus(
                    "Không đọc được thiết lập: " +
                    ex.Message,
                    true);
            }
        }

        private void ApplySnapshot(
            WorkbookEstimateV2SettingsSnapshot value)
        {
            if (value == null)
                return;

            workbookValue.Text =
                SafeWorkbookDisplayName();
            workbookDetail.Text =
                "Đã nhận diện";

            roleValue.Text =
                value.ConfiguredSheetCount
                    .ToString() +
                "/6";
            roleDetail.Text =
                value.ConfiguredSheetCount == 6
                    ? "Đã cấu hình"
                    : "Chưa đủ";

            metadataValue.Text =
                value.MetadataValid
                    ? "Đầy đủ"
                    : "Chưa đủ";
            metadataDetail.Text =
                value.MetadataValid
                    ? "Hợp lệ"
                    : "Cần đăng ký";

            autosaveValue.Text =
                value.Settings.AutoSaveEnabled
                    ? "Đang bật"
                    : "Đang tắt";
            autosaveDetail.Text =
                value.Settings.AutoSaveEnabled
                    ? "Mỗi " +
                        value.Settings
                            .AutoSaveMinutes +
                        " phút"
                    : "Không tự lưu";

            autoRestoreToggle.Checked =
                value.Settings
                    .AutoRestoreNormDisplay;
            validateOnOpenToggle.Checked =
                value.Settings
                    .ValidateOnOpen;
            formulaLinkToggle.Checked = true;
            autoSyncToggle.Checked =
                value.Settings.AutoSyncRows;
            hiddenColumnsToggle.Checked = true;
            warnMappingToggle.Checked =
                value.Settings
                    .WarnOnMappingLoss;

            foreach (KeyValuePair<
                EstimateV2OutputSlot,
                ComboBox> pair in
                mappingCombos)
            {
                PopulateCombo(
                    pair.Value,
                    value.Sheets,
                    value.SheetKeys.ContainsKey(
                        pair.Key)
                        ? value.SheetKeys[pair.Key]
                        : string.Empty);
            }
        }

        private static void PopulateCombo(
            ComboBox combo,
            IEnumerable<WorkbookSheetDescriptor> sheets,
            string selectedKey)
        {
            combo.BeginUpdate();
            try
            {
                combo.Items.Clear();
                combo.Items.Add(
                    new SheetOption(
                        string.Empty,
                        "(Chưa chọn)"));

                foreach (WorkbookSheetDescriptor sheet in
                    (sheets ??
                        Enumerable.Empty<
                            WorkbookSheetDescriptor>())
                        .OrderBy(
                            item => item.Name,
                            StringComparer
                                .CurrentCultureIgnoreCase))
                {
                    combo.Items.Add(
                        new SheetOption(
                            sheet.Key,
                            sheet.Name));
                }

                int selected = 0;
                for (int index = 1;
                    index < combo.Items.Count;
                    index++)
                {
                    SheetOption option =
                        combo.Items[index]
                            as SheetOption;
                    if (option != null &&
                        string.Equals(
                            option.Key,
                            selectedKey,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        selected = index;
                        break;
                    }
                }
                combo.SelectedIndex = selected;
            }
            finally
            {
                combo.EndUpdate();
            }
        }

        private void SaveSettings()
        {
            try
            {
                EstimateV2Settings baseSettings =
                    snapshot == null
                        ? EstimateV2SettingsPolicy
                            .Defaults()
                        : snapshot.Settings;

                var settings =
                    new EstimateV2Settings(
                        autoRestoreToggle.Checked,
                        validateOnOpenToggle.Checked,
                        true,
                        autoSyncToggle.Checked,
                        true,
                        true,
                        warnMappingToggle.Checked,
                        baseSettings.AutoSaveEnabled,
                        baseSettings.AutoSaveMinutes);

                var mappings =
                    new Dictionary<
                        EstimateV2OutputSlot,
                        string>();
                foreach (KeyValuePair<
                    EstimateV2OutputSlot,
                    ComboBox> pair in
                    mappingCombos)
                {
                    SheetOption option =
                        pair.Value.SelectedItem
                            as SheetOption;
                    mappings[pair.Key] =
                        option == null
                            ? string.Empty
                            : option.Key;
                }

                WorkbookEstimateV2SettingsService
                    .SaveConfiguration(
                        workbook,
                        settings,
                        mappings);

                EstimateV2Settings normalized =
                    WorkbookEstimateV2SettingsService
                        .Load(workbook);
                settingsChangedAction?.Invoke(
                    normalized);
                LoadSnapshot();
                ShowStatus(
                    "Đã lưu thiết lập. Mapping sheet dùng identity bền vững; công thức, Custom XML và cột kỹ thuật ẩn luôn được bảo vệ.",
                    false);
            }
            catch (Exception ex)
            {
                ShowStatus(
                    "Không lưu được thiết lập: " +
                    ex.Message,
                    true);
            }
        }

        private void ResetSettings()
        {
            try
            {
                WorkbookEstimateV2SettingsService
                    .Reset(workbook);
                EstimateV2Settings defaults =
                    EstimateV2SettingsPolicy.Defaults();
                settingsChangedAction?.Invoke(
                    defaults);
                LoadSnapshot();
                ShowStatus(
                    "Đã khôi phục hành vi mặc định. Mapping sheet hiện tại được giữ nguyên.",
                    false);
            }
            catch (Exception ex)
            {
                ShowStatus(
                    "Không khôi phục được mặc định: " +
                    ex.Message,
                    true);
            }
        }

        private void OpenConfigurationFolder()
        {
            try
            {
                AppPaths.EnsureRoot();
                string folder =
                    AppPaths.LocalAppDataRoot;
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                Process.Start(
                    new ProcessStartInfo
                    {
                        FileName = folder,
                        UseShellExecute = true
                    });
                ShowStatus(
                    "Đã mở thư mục cấu hình TTBMVN.",
                    false);
            }
            catch (Exception ex)
            {
                ShowStatus(
                    "Không mở được thư mục cấu hình: " +
                    ex.Message,
                    true);
            }
        }

        private string SafeWorkbookDisplayName()
        {
            try
            {
                string name =
                    workbook.Name ??
                    "Workbook";
                if (name.Length <= 15)
                    return name;
                return name.Substring(
                    0,
                    12) +
                    "...";
            }
            catch
            {
                return "Workbook";
            }
        }

        private void ShowStatus(
            string text,
            bool warning)
        {
            statusLabel.Text =
                text ?? string.Empty;
            statusLabel.BackColor =
                warning
                    ? Color.FromArgb(
                        255, 248, 226)
                    : GreenSoft;
            statusLabel.ForeColor =
                warning
                    ? Color.DarkGoldenrod
                    : GreenDark;
        }

        private static void PaintBorder(
            object sender,
            PaintEventArgs e)
        {
            Control control =
                (Control)sender;
            using (var pen =
                new Pen(Border))
            {
                e.Graphics.DrawRectangle(
                    pen,
                    0,
                    0,
                    Math.Max(
                        0,
                        control.Width - 1),
                    Math.Max(
                        0,
                        control.Height - 1));
            }
        }

        private sealed class MetricLabels
        {
            internal MetricLabels(
                Label value,
                Label detail)
            {
                Value = value;
                Detail = detail;
            }

            internal Label Value { get; }
            internal Label Detail { get; }
        }

        private sealed class SheetOption
        {
            internal SheetOption(
                string key,
                string name)
            {
                Key =
                    (key ?? string.Empty)
                        .Trim();
                Name =
                    (name ?? string.Empty)
                        .Trim();
            }

            internal string Key { get; }
            internal string Name { get; }

            public override string ToString()
            {
                return Name;
            }
        }

        private sealed class ToggleSwitch : CheckBox
        {
            internal bool Locked { get; set; }

            internal ToggleSwitch()
            {
                SetStyle(
                    ControlStyles.UserPaint |
                    ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer,
                    true);
                Cursor = Cursors.Hand;
            }

            protected override void OnClick(
                EventArgs e)
            {
                if (!Locked)
                    base.OnClick(e);
            }

            protected override void OnPaint(
                PaintEventArgs e)
            {
                e.Graphics.SmoothingMode =
                    SmoothingMode.AntiAlias;
                e.Graphics.Clear(
                    Parent == null
                        ? Color.White
                        : Parent.BackColor);

                Rectangle track =
                    new Rectangle(
                        1,
                        3,
                        Width - 3,
                        Height - 7);
                int radius =
                    Math.Max(
                        4,
                        track.Height / 2);

                using (GraphicsPath path =
                    RoundedRectangle(
                        track,
                        radius))
                using (var brush =
                    new SolidBrush(
                        Checked
                            ? Green
                            : Color.FromArgb(
                                185, 192, 198)))
                {
                    e.Graphics.FillPath(
                        brush,
                        path);
                }

                int knob =
                    Math.Max(
                        10,
                        track.Height - 4);
                int x =
                    Checked
                        ? track.Right -
                            knob - 2
                        : track.Left + 2;
                int y =
                    track.Top +
                    (track.Height -
                     knob) / 2;
                using (var knobBrush =
                    new SolidBrush(
                        Color.White))
                {
                    e.Graphics.FillEllipse(
                        knobBrush,
                        x,
                        y,
                        knob,
                        knob);
                }
            }

            private static GraphicsPath RoundedRectangle(
                Rectangle rect,
                int radius)
            {
                int diameter =
                    Math.Max(
                        2,
                        radius * 2);
                var path =
                    new GraphicsPath();
                path.AddArc(
                    rect.Left,
                    rect.Top,
                    diameter,
                    diameter,
                    180,
                    90);
                path.AddArc(
                    rect.Right -
                        diameter,
                    rect.Top,
                    diameter,
                    diameter,
                    270,
                    90);
                path.AddArc(
                    rect.Right -
                        diameter,
                    rect.Bottom -
                        diameter,
                    diameter,
                    diameter,
                    0,
                    90);
                path.AddArc(
                    rect.Left,
                    rect.Bottom -
                        diameter,
                    diameter,
                    diameter,
                    90,
                    90);
                path.CloseFigure();
                return path;
            }
        }
    }
}
