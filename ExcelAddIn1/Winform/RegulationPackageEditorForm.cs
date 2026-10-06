using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelAddIn1.Winform
{
    internal sealed class RegulationPackageEditorForm : Form
    {
        private readonly TextBox id = new TextBox(), version = new TextBox(), note = new TextBox();
        private readonly DateTimePicker from = new DateTimePicker(), to = new DateTimePicker();
        private readonly Dictionary<RegulationModuleKind, DataGridView> grids = new Dictionary<RegulationModuleKind, DataGridView>();
        private readonly DataGridView sources;
        private readonly TabControl tabs;
        private readonly Label status;
        private bool loading = true, dirty;
        private string checkedSource, checkedBundle;
        internal RegulationPackage InstalledPackage { get; private set; }

        internal RegulationPackageEditorForm(RegulationPackageBundle template, bool create)
        {
            Text = create ? "Tạo gói pháp lý từ mẫu" : "Sửa gói pháp lý - phiên bản mới";
            StartPosition = FormStartPosition.CenterParent; Size = new Size(1100, 720);
            MinimumSize = new Size(760, 500); Font = new Font("Segoe UI", 9f);
            BackColor = Color.White; ForeColor = Color.FromArgb(33, 43, 54);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(12) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);
            var fields = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 4 };
            for (int i = 0; i < 4; i++) fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            AddField(fields, "Mã gói", id, 0); AddField(fields, "Phiên bản", version, 1);
            from.Format = to.Format = DateTimePickerFormat.Custom;
            from.CustomFormat = to.CustomFormat = "dd/MM/yyyy"; to.ShowCheckBox = true;
            AddField(fields, "Hiệu lực từ", from, 2); AddField(fields, "Hiệu lực đến", to, 3);
            fields.Controls.Add(new Label { Text = "Ghi chú / căn cứ thay đổi", AutoSize = true }, 0, 2);
            note.Dock = DockStyle.Fill; fields.Controls.Add(note, 0, 3); fields.SetColumnSpan(note, 4);
            root.Controls.Add(fields, 0, 0);
            tabs = new TabControl { Dock = DockStyle.Fill };
            root.Controls.Add(tabs, 0, 1);
            foreach (RegulationModuleKind kind in Enum.GetValues(typeof(RegulationModuleKind)))
            {
                var grid = NewGrid();
                string[] names = { "Mã", "Loại bản ghi", "Đơn vị", "Tên dữ liệu", "Biến thể / hao phí / quy tắc", "Mã văn bản", "Trang từ", "Trang đến", "Mục căn cứ" };
                foreach (string name in names) grid.Columns.Add(name, name);
                grid.Columns.Add(new DataGridViewComboBoxColumn { HeaderText = "Kiểm chứng", Name = "Verification",
                    DisplayMember = "Text", ValueMember = "Value", DataSource = new[] {
                        new { Text = "Chưa kiểm chứng", Value = "Unverified" },
                        new { Text = "Đã đối chiếu văn bản gốc", Value = "VerifiedAgainstOfficialSource" },
                        new { Text = "Suy dẫn từ dữ liệu đã kiểm chứng", Value = "DerivedFromVerifiedRecords" } } });
                grid.Columns[3].Width = 260; grid.Columns[4].Width = 400; grid.Columns[9].Width = 240;
                var defaults = template.Modules[kind].Records.First();
                grid.DefaultValuesNeeded += (s, e) => {
                    e.Row.Cells[1].Value = defaults.RecordType; e.Row.Cells[2].Value = defaults.Unit;
                    e.Row.Cells[5].Value = defaults.Source.DocumentId;
                    e.Row.Cells[6].Value = e.Row.Cells[7].Value = 1;
                    e.Row.Cells[9].Value = "Unverified";
                };
                foreach (var record in template.Modules[kind].Records)
                    grid.Rows.Add(record.Key, record.RecordType, record.Unit, record.Title, record.Data,
                        record.Source.DocumentId, record.Source.PageFrom, record.Source.PageTo, record.Source.Section, record.Verification.ToString());
                grid.CellValueChanged += (s, e) =>
                {
                    if (loading || e.RowIndex < 0) return;
                    if (e.ColumnIndex != 9) grid.Rows[e.RowIndex].Cells[9].Value = "Unverified";
                    Changed();
                };
                grid.UserDeletedRow += (s, e) => Changed();
                grids.Add(kind, grid);
                string title = kind == RegulationModuleKind.Norm ? "Định mức" : kind == RegulationModuleKind.CostRule ? "Chi phí" :
                    kind == RegulationModuleKind.MachineRate ? "Giá máy" : kind == RegulationModuleKind.Geography ? "Địa bàn" :
                    kind == RegulationModuleKind.Compliance ? "Tuân thủ" : "Quy trình";
                var page = new TabPage(title); page.Controls.Add(grid); tabs.TabPages.Add(page);
            }
            sources = NewGrid();
            foreach (string name in new[] { "Mã văn bản", "Tên văn bản", "Cơ quan", "Ngày ban hành", "Hiệu lực từ", "Hiệu lực đến", "URL chính thức", "SHA-256 tài liệu" })
                sources.Columns.Add(name, name);
            sources.Columns[1].Width = 300; sources.Columns[6].Width = 300; sources.Columns[7].Width = 450;
            sources.Columns[3].HeaderText = "Ngày ban hành (yyyy-MM-dd)";
            sources.Columns[4].HeaderText = "Hiệu lực từ (yyyy-MM-dd)";
            sources.Columns[5].HeaderText = "Hiệu lực đến (yyyy-MM-dd)";
            foreach (var source in template.Package.Sources)
                sources.Rows.Add(source.DocumentId, source.Title, source.Publisher, Date(source.IssuedDate), Date(source.EffectiveFrom),
                    Date(source.EffectiveTo), source.OfficialUri, source.ContentChecksum);
            sources.CellValueChanged += (s, e) => { if (!loading) { foreach (var grid in grids.Values)
                foreach (DataGridViewRow row in grid.Rows) if (!row.IsNewRow) row.Cells[9].Value = "Unverified"; Changed(); } };
            sources.UserDeletedRow += (s, e) => Changed();
            var sourcePage = new TabPage("Văn bản nguồn"); sourcePage.Controls.Add(sources); tabs.TabPages.Add(sourcePage);
            status = new Label { Dock = DockStyle.Fill, ForeColor = Color.FromArgb(35, 88, 180), AutoEllipsis = true };
            root.Controls.Add(status, 0, 2);
            var commands = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            root.Controls.Add(commands, 0, 3);
            Button(commands, "Thêm dòng", () => { var grid = CurrentGrid(); grid.CurrentCell = grid.Rows[grid.NewRowIndex].Cells[0]; grid.BeginEdit(true); });
            Button(commands, "Xóa dòng", DeleteRows);
            Button(commands, "Lưu bản nháp", () => { checkedSource = SaveDraft(); status.Text = "Đã lưu nháp: " + checkedSource; dirty = false; });
            Button(commands, "Mở bản nháp", OpenDraft);
            Button(commands, "Kiểm tra gói", CheckPackage);
            Button(commands, "Cài gói", Install);
            Button(commands, "Đóng", Close);
            id.Text = create ? "TTBMVN-CUSTOM-" + DateTime.Now.ToString("yyyyMMddHHmmss") : template.Package.PackageId;
            Version current; version.Text = !create && Version.TryParse(template.Package.DataVersion, out current)
                ? current.Major + "." + current.Minor + "." + (current.Build + 1) : "1.0.0";
            from.Value = template.Package.EffectiveFrom; to.Value = template.Package.EffectiveTo ?? DateTime.Today;
            to.Checked = template.Package.EffectiveTo.HasValue;
            note.Text = "Gói tùy chỉnh từ " + template.Package.PackageId + " @ " + template.Package.DataVersion;
            id.TextChanged += (s, e) => Changed(); version.TextChanged += (s, e) => Changed(); note.TextChanged += (s, e) => Changed();
            from.ValueChanged += (s, e) => Changed(); to.ValueChanged += (s, e) => Changed();
            FormClosing += (s, e) => {
                if (!dirty) return;
                var choice = MessageBox.Show(this, "Lưu bản nháp trước khi đóng?", "Gói pháp lý",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                e.Cancel = choice == DialogResult.Cancel;
                if (choice == DialogResult.Yes) { e.Cancel = true; Try(() => { SaveDraft(); dirty = false; e.Cancel = false; }); }
            };
            loading = false; dirty = true;
            status.Text = "Gói tùy chỉnh; chưa thay đổi workbook hoặc gói mẫu.";
        }

        private static void AddField(TableLayoutPanel table, string text, Control control, int column)
        { table.Controls.Add(new Label { Text = text, AutoSize = true }, column, 0); control.Dock = DockStyle.Fill; table.Controls.Add(control, column, 1); }
        private DataGridView NewGrid()
        {
            var grid = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, ForeColor = ForeColor,
                AutoGenerateColumns = false, AllowUserToAddRows = true, AllowUserToDeleteRows = true, MultiSelect = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect };
            grid.DataError += (s, e) => { e.ThrowException = false; if (status != null) status.Text = "Giá trị kiểm chứng không hợp lệ."; };
            return grid;
        }
        private void Button(FlowLayoutPanel parent, string text, Action action)
        { var button = new Button { Text = text, AutoSize = true, Height = 32 }; button.Click += (s, e) => Try(action); parent.Controls.Add(button); }
        private void Try(Action action)
        { try { action(); } catch (Exception ex) { RuntimeLogger.Log(ex, "Package editor"); status.Text = ex.Message; MessageBox.Show(this, ex.Message, "Gói chưa hợp lệ", MessageBoxButtons.OK, MessageBoxIcon.Warning); } }
        private void Changed() { if (loading) return; dirty = true; checkedBundle = null; }
        private DataGridView CurrentGrid() => (DataGridView)tabs.SelectedTab.Controls[0];
        private void DeleteRows()
        {
            var grid = CurrentGrid(); var rows = grid.SelectedRows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).ToArray();
            if (rows.Length == 0 || MessageBox.Show(this, "Xóa " + rows.Length + " dòng trong bản nháp?", "Xóa dữ liệu",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            foreach (var row in rows) grid.Rows.Remove(row); Changed();
        }
        private void EndEdit() { foreach (var grid in grids.Values) grid.EndEdit(); sources.EndEdit(); Validate(); }
        private string SaveDraft()
        {
            EndEdit();
            var records = grids.ToDictionary(p => p.Key, p => (IReadOnlyList<RegulationDataRecord>)p.Value.Rows.Cast<DataGridViewRow>()
                .Where(r => !r.IsNewRow).Select(r => new RegulationDataRecord(Cell(r, 0), Cell(r, 1), Cell(r, 2), Cell(r, 3), Cell(r, 4),
                    new RegulationSourceLocator(Cell(r, 5), int.Parse(Cell(r, 6), CultureInfo.InvariantCulture), int.Parse(Cell(r, 7), CultureInfo.InvariantCulture), Cell(r, 8)),
                    (RegulationDataVerification)Enum.Parse(typeof(RegulationDataVerification), Cell(r, 9)))).ToArray());
            var citations = sources.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => new RegulationPackageSourceDocument(
                Cell(r, 0), Cell(r, 1), Cell(r, 2), ParseDate(Cell(r, 3)), ParseDate(Cell(r, 4)),
                string.IsNullOrWhiteSpace(Cell(r, 5)) ? (DateTime?)null : ParseDate(Cell(r, 5)), Cell(r, 6), Cell(r, 7))).ToArray();
            string source = Path.Combine(AppPaths.LocalAppDataRoot, "PackageDrafts", Guid.NewGuid().ToString("N"));
            RegulationPackageAuthoring.WriteSource(source, id.Text.Trim(), version.Text.Trim(), from.Value.Date,
                to.Checked ? to.Value.Date : (DateTime?)null, note.Text, citations, records);
            return source;
        }
        private void CheckPackage()
        {
            checkedBundle = null; checkedSource = SaveDraft();
            string output = Path.Combine(AppPaths.LocalAppDataRoot, "PackageBuilds", Guid.NewGuid().ToString("N"));
            var bundle = RegulationPackageAuthoring.BuildChecked(checkedSource, output);
            checkedBundle = bundle.Directory; status.Text = "Kiểm tra cấu trúc đạt. SHA-256: " + bundle.Package.PackageChecksum;
        }
        private void Install()
        {
            EndEdit(); CheckPackage();
            if (MessageBox.Show(this, "Cài gói tùy chỉnh " + id.Text + " @ " + version.Text + "?\nKiểm tra cấu trúc không thay thế kiểm chứng pháp lý. Workbook chưa đổi gói.",
                "Cài gói", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            InstalledPackage = new RegulationPackageStore(AppPaths.RegulationPackageDirectory).ImportFromDirectory(checkedBundle).Package;
            dirty = false; DialogResult = DialogResult.OK; Close();
        }
        private void OpenDraft()
        {
            if (dirty && MessageBox.Show(this, "Thay nội dung chưa lưu bằng bản nháp đã chọn?", "Mở bản nháp", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            using (var dialog = new FolderBrowserDialog { Description = "Chọn bản nháp chứa package.properties", SelectedPath = Path.Combine(AppPaths.LocalAppDataRoot, "PackageDrafts") })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                var draft = RegulationPackageSourceReader.Read(dialog.SelectedPath);
                loading = true;
                try
                {
                    id.Text = draft.PackageId; version.Text = draft.DataVersion; note.Text = draft.TransitionNote;
                    from.Value = draft.EffectiveFrom; to.Value = draft.EffectiveTo ?? DateTime.Today; to.Checked = draft.EffectiveTo.HasValue;
                    foreach (var pair in grids) { pair.Value.Rows.Clear(); foreach (var r in draft.Records[pair.Key])
                        pair.Value.Rows.Add(r.Key, r.RecordType, r.Unit, r.Title, r.Data, r.Source.DocumentId, r.Source.PageFrom, r.Source.PageTo, r.Source.Section, r.Verification.ToString()); }
                    sources.Rows.Clear(); foreach (var s in draft.Sources)
                        sources.Rows.Add(s.DocumentId, s.Title, s.Publisher, Date(s.IssuedDate), Date(s.EffectiveFrom), Date(s.EffectiveTo), s.OfficialUri, s.ContentChecksum);
                }
                finally { loading = false; }
                dirty = false; checkedBundle = null; status.Text = "Đã mở nháp: " + dialog.SelectedPath;
            }
        }
        private static string Cell(DataGridViewRow row, int column) => Convert.ToString(row.Cells[column].Value, CultureInfo.InvariantCulture) ?? "";
        private static string Date(DateTime? value) => value?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
        private static DateTime ParseDate(string text) => DateTime.ParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture);
    }
}
