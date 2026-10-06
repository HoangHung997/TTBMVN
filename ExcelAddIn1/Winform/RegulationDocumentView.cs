using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace ExcelAddIn1.Winform
{
    internal sealed class RegulationDocumentView : UserControl
    {
        private readonly Func<IEnumerable<DataGridViewRow>> getRows;
        private readonly Func<DataGridView> getNewGrid;
        private readonly ComboBox selector = new ComboBox();
        private readonly DataGridView metadata = Grid(), quantities = Grid(), parameters = Grid();
        private readonly TextBox variants = new TextBox();
        private readonly Label heading = new Label();
        private readonly TabControl details = new TabControl();
        private readonly Label message = new Label();
        private DataGridViewRow current;
        private string original;
        private bool loading, norm;
        internal event Action Modified;
        private readonly CultureInfo numberCulture = ExcelAddIn1.Funtion.ExcelCulture.GetNumberCulture(Globals.ThisAddIn.Application);

        internal RegulationDocumentView(Func<IEnumerable<DataGridViewRow>> rows, Func<DataGridView> newGrid)
        {
            getRows = rows; getNewGrid = newGrid; Dock = DockStyle.Fill; BackColor = Color.White;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(8) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); Controls.Add(root);
            selector.Dock = DockStyle.Fill; selector.DropDownStyle = ComboBoxStyle.DropDownList; selector.DropDownWidth = 850;
            root.Controls.Add(selector, 0, 0);
            heading.Dock = DockStyle.Fill; heading.Font = new Font("Segoe UI", 11f, FontStyle.Bold);
            heading.AutoEllipsis = false; root.Controls.Add(heading, 0, 1);
            details.Dock = DockStyle.Fill; root.Controls.Add(details, 0, 2);
            var ratePage = new TabPage("Bảng hao phí"); ratePage.Controls.Add(quantities); details.TabPages.Add(ratePage);
            var parameterPage = new TabPage("Điều kiện / thông số"); parameterPage.Controls.Add(parameters); details.TabPages.Add(parameterPage);
            variants.Dock = DockStyle.Top; variants.Height = 28; parameterPage.Controls.Add(variants);
            metadata.Columns.Add("Field", "Thông tin / căn cứ"); metadata.Columns.Add("Value", "Nội dung");
            metadata.Columns[0].ReadOnly = true; metadata.Columns[0].Width = 180; metadata.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            metadata.AllowUserToAddRows = metadata.AllowUserToDeleteRows = false;
            var metadataPage = new TabPage("Thông tin / căn cứ"); metadataPage.Controls.Add(metadata); details.TabPages.Add(metadataPage);
            parameters.Columns.Add("Key", "Thông số"); parameters.Columns.Add("Value", "Giá trị / điều kiện");
            parameters.Columns[0].Width = 230; parameters.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            parameters.CellFormatting += (s, e) => { if (e.ColumnIndex == 0 && e.Value is string key) {
                e.Value = ParameterTitle(key); e.FormattingApplied = true; } };
            message.Dock = DockStyle.Fill; message.ForeColor = Color.FromArgb(180, 75, 20); root.Controls.Add(message, 0, 3);
            foreach (var grid in new[] { metadata, quantities, parameters }) {
                grid.CellValueChanged += (s, e) => { if (!loading) Modified?.Invoke(); };
                grid.CellBeginEdit += (s, e) => { if (!loading) Modified?.Invoke(); };
                grid.UserDeletedRow += (s, e) => { if (!loading) Modified?.Invoke(); };
                grid.DataError += (s, e) => { e.ThrowException = false; message.Text = "Giá trị chưa hợp lệ."; };
            }
            variants.TextChanged += (s, e) => { if (!loading) Modified?.Invoke(); };
            selector.SelectedIndexChanged += (s, e) => {
                if (loading) return;
                var next = selector.SelectedItem as Item;
                try { Flush(); LoadRecord(next?.Row); }
                catch (Exception ex) { message.Text = ex.Message; loading = true;
                    selector.SelectedItem = selector.Items.Cast<Item>().FirstOrDefault(i => i.Row == current); loading = false; }
            };
            RefreshRecords();
        }

        internal void RefreshRecords()
        {
            loading = true;
            try {
                var previous = current; selector.Items.Clear();
                foreach (var row in getRows().Where(r => !r.IsNewRow)) selector.Items.Add(new Item(row));
                selector.SelectedItem = selector.Items.Cast<Item>().FirstOrDefault(i => i.Row == previous);
                if (selector.SelectedIndex < 0 && selector.Items.Count > 0) selector.SelectedIndex = 0;
                LoadRecord((selector.SelectedItem as Item)?.Row);
            } finally { loading = false; }
        }

        internal void Flush()
        {
            if (current == null) return;
            quantities.EndEdit(); parameters.EndEdit(); metadata.EndEdit();
            string data;
            var fields = parameters.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow)
                .Select(r => new KeyValuePair<string, string>(CellText(r, 0), CellText(r, 1))).ToArray();
            if (norm) {
                var all = RegulationRecordTable.ReadFields(original);
                foreach (var key in all.Keys.Where(k => k != "variantCodes" && k != "rates").ToArray()) all.Remove(key);
                foreach (var field in fields) all.Add(field.Key, field.Value);
                var rows = quantities.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r =>
                    r.Cells.Cast<DataGridViewCell>().Select(c => Convert.ToString(c.Value, CultureInfo.InvariantCulture) ?? "").ToArray()).ToArray();
                foreach (var row in rows) for (int c = 3; c < row.Length; c++) {
                    decimal value;
                    if (!decimal.TryParse(row[c], NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, numberCulture, out value))
                        throw new FormatException("Hao phí không hợp lệ theo dấu thập phân Excel.");
                    row[c] = value.ToString(CultureInfo.InvariantCulture);
                }
                data = RegulationRecordTable.WriteNorm(RegulationRecordTable.WriteFields(all), variants.Text.Split(','), rows);
            } else data = RegulationRecordTable.WriteFields(fields);
            string[] values = metadata.Rows.Cast<DataGridViewRow>().Select(r => CellText(r, 1)).ToArray();
            if (CellText(current, 9) != values[8]) current.Cells[9].Value = values[8];
            for (int i = 0; i < MetadataColumns.Length - 1; i++) {
                int column = MetadataColumns[i];
                if (CellText(current, column) != values[i]) current.Cells[column].Value = values[i];
            }
            if (CellText(current, 4) != data) current.Cells[4].Value = data;
            metadata.Rows[8].Cells[1].Value = CellText(current, 9);
            original = data; heading.Text = DisplayCode(CellText(current, 0)) + " — " + CellText(current, 3) + "\r\nĐơn vị: " + UnitTitle(CellText(current, 2));
            message.Text = "";
        }

        internal void AddRecord()
        {
            Flush(); var grid = current?.DataGridView ?? getNewGrid();
            var sample = current ?? grid.Rows.Cast<DataGridViewRow>().FirstOrDefault(r => !r.IsNewRow);
            if (sample == null) throw new InvalidOperationException("Không có bản ghi mẫu.");
            int index = grid.Rows.Add(sample.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray());
            grid.Rows[index].Cells[0].Value = CellText(sample, 0) + "-NEW-" + Guid.NewGuid().ToString("N").Substring(0, 6);
            grid.Rows[index].Cells[3].Value = "Công tác mới"; grid.Rows[index].Cells[9].Value = "Unverified";
            current = grid.Rows[index]; RefreshRecords();
        }

        internal void DeleteRecord()
        {
            if (current == null || MessageBox.Show(this, "Xóa công tác và toàn bộ bảng chi tiết trong bản nháp?", "Xóa công tác",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            current.DataGridView.Rows.Remove(current); current = null; RefreshRecords();
        }

        private static readonly int[] MetadataColumns = { 0, 3, 2, 1, 5, 6, 7, 8, 9 };
        private static readonly string[] MetadataLabels = { "Mã công tác", "Tên công tác", "Đơn vị", "Loại dữ liệu", "Văn bản", "Trang từ", "Trang đến", "Mục căn cứ", "Kiểm chứng" };
        private void LoadRecord(DataGridViewRow row)
        {
            bool previous = loading; loading = true;
            try { PopulateRecord(row); } finally { loading = previous; }
        }
        private void PopulateRecord(DataGridViewRow row)
        {
            current = row; metadata.Rows.Clear(); parameters.Rows.Clear(); quantities.Columns.Clear();
            heading.Text = ""; message.Text = "";
            if (row == null) return;
            original = CellText(row, 4);
            var fields = RegulationRecordTable.ReadFields(original);
            norm = CellText(row, 1) == "NormCatalog" && fields.ContainsKey("variantCodes") && fields.ContainsKey("rates");
            for (int i = 0; i < MetadataColumns.Length; i++) metadata.Rows.Add(MetadataLabels[i], CellText(row, MetadataColumns[i]));
            var verification = new DataGridViewComboBoxCell { DisplayMember = "Label", ValueMember = "Code", DataSource = new[] {
                new { Label = "Chưa kiểm chứng", Code = "Unverified" },
                new { Label = "Đã đối chiếu văn bản gốc", Code = "VerifiedAgainstOfficialSource" },
                new { Label = "Suy dẫn từ dữ liệu kiểm chứng", Code = "DerivedFromVerifiedRecords" } } };
            verification.Value = CellText(row, 9); metadata.Rows[8].Cells[1] = verification;
            metadata.Rows[3].Cells[1].ReadOnly = true;
            heading.Text = DisplayCode(CellText(row, 0)) + " — " + CellText(row, 3) + "\r\nĐơn vị: " + UnitTitle(CellText(row, 2));
            variants.Visible = norm;
            if (norm) {
                var record = new RegulationDataRecord(CellText(row, 0), CellText(row, 1), CellText(row, 2), CellText(row, 3), original,
                    new RegulationSourceLocator(CellText(row, 5), int.Parse(CellText(row, 6)), int.Parse(CellText(row, 7)), CellText(row, 8)), RegulationDataVerification.Unverified);
                var definition = RegulationRecordTable.ReadNorm(record);
                variants.Text = string.Join(",", definition.Variants);
                quantities.Columns.Add(new DataGridViewComboBoxColumn { Name = "Kind", HeaderText = "Loại hao phí", DisplayMember = "Label", ValueMember = "Code", DataSource = new[] {
                    new { Label = "Vật liệu", Code = "Material" }, new { Label = "Nhân công", Code = "Labor" }, new { Label = "Máy", Code = "Machine" } } });
                quantities.Columns.Add("Resource", "Mã vật tư / nhân công / máy"); quantities.Columns.Add("Unit", "Đơn vị");
                quantities.Columns[1].Width = 220;
                for (int i = 0; i < definition.Variants.Count; i++) {
                    int c = quantities.Columns.Add("V" + i, DisplayCode(definition.Key) + "." + (i + 1) + "\n" + VariantTitle(definition.Variants[i]));
                    quantities.Columns[c].Width = 150; quantities.Columns[c].ToolTipText = definition.Variants[i];
                }
                foreach (var rate in definition.Rates) quantities.Rows.Add(new object[] { rate.Kind.ToString(), rate.ResourceCode, rate.Unit }
                    .Concat(rate.Quantities.Select(q => (object)q.ToString(numberCulture))).ToArray());
                foreach (var field in fields.Where(p => p.Key != "variantCodes" && p.Key != "rates")) parameters.Rows.Add(field.Key, field.Value);
                details.SelectedIndex = 0;
            } else {
                foreach (var field in fields) parameters.Rows.Add(field.Key, field.Value);
                details.SelectedIndex = 1;
            }
            details.TabPages[0].Enabled = norm;
        }

        private static string VariantTitle(string key)
        {
            if (key.StartsWith("density-")) return "Mật độ " + key.Substring(8);
            if (key.StartsWith("forest-")) return "Rừng loại " + key.Substring(7);
            if (key.StartsWith("soil-")) return "Đất cấp " + key.Substring(5);
            if (key.StartsWith("water-")) return "Nước sâu " + key.Substring(6) + " m";
            if (key.StartsWith("depth-")) return "Độ sâu " + key.Substring(6) + " m";
            if (key == "standard") return "Thông thường";
            if (key == "plain-midland") return "Đồng bằng, trung du";
            if (key == "mountain-island") return "Miền núi, hải đảo";
            return key;
        }
        private static string ParameterTitle(string key)
        {
            switch (key) {
                case "zone": return "Khu vực";
                case "funding": return "Nguồn kinh phí";
                case "landRate": return "Suất dự toán trên cạn";
                case "underwaterTo12mRate": return "Suất dự toán dưới nước đến 12 m";
                case "priority": return "Thứ tự ưu tiên";
                case "scope": return "Phạm vi áp dụng";
                case "province": return "Tỉnh / thành phố";
                case "zones": return "Các khu vực";
                case "fallbackZone": return "Khu vực mặc định";
                case "tableRow": return "Dòng trong bảng văn bản";
                case "article": return "Điều";
                case "chapter": return "Chương";
                case "group": return "Nhóm công tác";
                case "variants": return "Số loại";
                case "forestClasses": return "Số loại rừng";
                case "adjustments": return "Điều chỉnh hao phí";
                case "constraints": return "Ràng buộc áp dụng";
                case "ratePercent": return "Tỷ lệ (%)";
                case "base": return "Cơ sở tính";
                case "code": return "Mã chi phí";
                case "rounding": return "Quy tắc làm tròn";
                case "annualShifts": return "Số ca trong năm";
                case "depreciationPercent": return "Tỷ lệ khấu hao (%)";
                case "repairPercent": return "Tỷ lệ sửa chữa (%)";
                case "otherPercent": return "Chi phí khác (%)";
                case "fuel": return "Nhiên liệu / năng lượng";
                case "operators": return "Nhân công điều khiển";
                case "referencePriceVnd": return "Nguyên giá tham chiếu (đồng)";
                case "resolvedDocument": return "Văn bản hợp nhất";
                case "baseStandard": return "Quy chuẩn gốc";
                case "amendedBy": return "Văn bản sửa đổi";
                case "resolution": return "Cách xác định khu vực";
                case "conflictRule": return "Quy tắc xử lý trùng khu vực";
                default: return key;
            }
        }
        private static string UnitTitle(string key)
        {
            switch (key) {
                case "commune": return "Xã";
                case "m2-10000": return "10.000 m²";
                case "signal": return "Tín hiệu";
                case "million-vnd-per-ha": return "Triệu đồng/ha";
                case "locality": return "Địa bàn";
                case "worker-day": return "Công";
                case "shift": return "Ca";
                default: return key;
            }
        }
        private static string DisplayCode(string key) => key.StartsWith("NORM-") ? key.Substring(5) : key;
        private static string CellText(DataGridViewRow r, int c) => Convert.ToString(r.Cells[c].Value, CultureInfo.InvariantCulture) ?? "";
        private static DataGridView Grid() => new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White,
            ForeColor = Color.FromArgb(33, 43, 54), AutoGenerateColumns = false, AllowUserToAddRows = true,
            AllowUserToDeleteRows = true, ColumnHeadersHeight = 52, ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells };
        private sealed class Item
        {
            internal readonly DataGridViewRow Row;
            internal Item(DataGridViewRow row) { Row = row; }
            public override string ToString() => DisplayCode(CellText(Row, 0)) + " — " + CellText(Row, 3);
        }
    }
}
