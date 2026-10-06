using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class ResultAuditControl : UserControl
    {
        private readonly Excel.Workbook workbook;
        private readonly Label locationLabel;
        private readonly Label identityLabel;
        private readonly Label normLabel;
        private readonly Label statusLabel;
        private readonly DataGridView sourceGrid;

        public ResultAuditControl(Excel.Workbook workbook)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            Padding = new Padding(14);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));

            var header = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            var refreshButton = new Button
            {
                Name = "btnReadSelectedAudit",
                Text = "Đọc ô đang chọn",
                Width = 132,
                Height = 30
            };
            refreshButton.Click += (sender, args) => RefreshSelection();
            header.Controls.Add(refreshButton);
            header.Controls.Add(new Label
            {
                Text = "Truy vết căn cứ kết quả",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(12, 7, 0, 0)
            });

            locationLabel = CreateInfoLabel();
            identityLabel = CreateInfoLabel();
            normLabel = CreateInfoLabel();
            statusLabel = CreateInfoLabel();
            statusLabel.ForeColor = Color.FromArgb(35, 92, 55);

            sourceGrid = new DataGridView
            {
                Name = "dgvAuditSources",
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            sourceGrid.Columns.Add(TextColumn("kind", "Loại", 86));
            sourceGrid.Columns.Add(TextColumn("document", "Văn bản", 150));
            sourceGrid.Columns.Add(TextColumn("pages", "Trang", 72));
            sourceGrid.Columns.Add(TextColumn("section", "Phụ lục / bảng / mục", 180));
            DataGridViewTextBoxColumn reference = TextColumn("reference", "Tham chiếu", 220);
            reference.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            sourceGrid.Columns.Add(reference);

            root.Controls.Add(header, 0, 0);
            root.Controls.Add(locationLabel, 0, 1);
            root.Controls.Add(identityLabel, 0, 2);
            root.Controls.Add(normLabel, 0, 3);
            root.Controls.Add(sourceGrid, 0, 4);
            root.Controls.Add(statusLabel, 0, 5);
            Controls.Add(root);

            Clear("Chọn một ô kết quả trong Excel rồi bấm Đọc ô đang chọn.");
        }

        public ResultAuditEntry CurrentEntry { get; private set; }

        public bool RefreshSelection()
        {
            try
            {
                ResultAuditEntry entry = WorkbookResultAuditService.FindAtActiveCell(workbook);
                if (entry == null)
                {
                    Clear("Ô đang chọn chưa có dữ liệu truy vết. Hãy ghi lại khối kết quả bằng app.");
                    return false;
                }
                ShowEntry(entry);
                return true;
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Read result audit at active cell", "DT-502", "lookup", workbook, string.Empty);
                Clear(ex.Message);
                return false;
            }
        }

        private void ShowEntry(ResultAuditEntry entry)
        {
            CurrentEntry = entry;
            locationLabel.Text = string.Format(
                CultureInfo.InvariantCulture,
                "{0} | {1} | R{2}C{3}:R{4}C{5}",
                entry.WorksheetRoleId,
                entry.Label,
                entry.FirstRow,
                entry.FirstColumn,
                entry.LastRow,
                entry.LastColumn);
            identityLabel.Text =
                "Package: " + entry.PackageId + " @ " + entry.PackageVersion + Environment.NewLine +
                "PriceProfile: " + (entry.PriceProfileId.Length == 0
                    ? "Không áp dụng"
                    : entry.PriceProfileId + " @ " + entry.PriceProfileVersion);
            normLabel.Text = entry.NormKey.Length == 0
                ? "Định mức: khối tổng hợp"
                : "Định mức: " + entry.NormKey +
                    (entry.VariantCode.Length == 0 ? string.Empty : " / " + entry.VariantCode);
            sourceGrid.Rows.Clear();
            ResultAuditSource[] visibleSources = entry.Sources
                .Where(source => source.Kind != ResultAuditSourceKind.Metadata)
                .ToArray();
            foreach (ResultAuditSource source in visibleSources)
            {
                sourceGrid.Rows.Add(
                    SourceKindText(source.Kind),
                    source.DocumentId,
                    PageText(source),
                    source.Section,
                    source.Reference);
            }
            statusLabel.Text = visibleSources.Length.ToString(CultureInfo.CurrentCulture) +
                " nguồn được lưu trong workbook; checksum package/profile đã xác thực khi đọc.";
        }

        private void Clear(string status)
        {
            CurrentEntry = null;
            locationLabel.Text = "Chưa có kết quả được chọn";
            identityLabel.Text = "Package: -" + Environment.NewLine + "PriceProfile: -";
            normLabel.Text = "Định mức: -";
            sourceGrid.Rows.Clear();
            statusLabel.Text = status ?? string.Empty;
        }

        private static Label CreateInfoLabel()
        {
            return new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static DataGridViewTextBoxColumn TextColumn(string name, string text, int width)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = text,
                Width = width,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private static string PageText(ResultAuditSource source)
        {
            if (source.PageFrom <= 0)
                return string.Empty;
            return source.PageTo > source.PageFrom
                ? source.PageFrom.ToString(CultureInfo.InvariantCulture) + "-" +
                    source.PageTo.ToString(CultureInfo.InvariantCulture)
                : source.PageFrom.ToString(CultureInfo.InvariantCulture);
        }

        private static string SourceKindText(ResultAuditSourceKind kind)
        {
            switch (kind)
            {
                case ResultAuditSourceKind.Regulation: return "Pháp lý";
                case ResultAuditSourceKind.Price: return "Giá";
                case ResultAuditSourceKind.Quantity: return "Khối lượng";
                case ResultAuditSourceKind.Output: return "Kết quả";
                case ResultAuditSourceKind.External: return "Căn cứ ngoài";
                case ResultAuditSourceKind.Metadata: return "Dữ liệu app";
                default: return kind.ToString();
            }
        }
    }
}
