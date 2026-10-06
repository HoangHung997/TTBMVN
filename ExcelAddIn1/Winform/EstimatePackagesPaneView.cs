using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal sealed class EstimatePackagesPaneView : EstimateActionPane
    {
        private readonly Excel.Workbook workbook;
        private readonly Label pinned;
        private readonly DataGridView packages;
        private readonly DateTimePicker preparedDate;
        private readonly DateTimePicker priceDate;
        private readonly TextBox impact;
        private readonly Button apply;
        private EstimateV2PackagePreview preview;

        internal EstimatePackagesPaneView(Excel.Workbook workbook, Action back)
            : base("Gói pháp lý & Dữ liệu", EstimateUiIconKind.Database, back)
        {
            this.workbook = workbook;
            Section("1. Gói đang pin trong workbook");
            pinned = new Label { AutoSize = true, Padding = new Padding(10),
                BackColor = Color.FromArgb(235, 248, 239), MinimumSize = new Size(0, 100) };
            Add(pinned);
            Section("2. Gói có sẵn trên máy");
            packages = new DataGridView { Height = 230, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                ReadOnly = true, MultiSelect = false, RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells };
            packages.Columns.Add("Id", "Mã gói"); packages.Columns.Add("Version", "Phiên bản");
            packages.Columns.Add("From", "Hiệu lực từ"); packages.Columns.Add("Norms", "Định mức");
            packages.SelectionChanged += (s, e) => InvalidatePreview();
            Add(packages);
            Add(Command("Tạo gói mới từ mẫu đã chọn", EstimateUiIconKind.Document, () => Run(() => EditPackage(true)), false));
            Add(Command("Sửa gói / định mức", EstimateUiIconKind.Database, () => Run(() => EditPackage(false)), false));
            Add(Command("Xóa phiên bản gói đã chọn", EstimateUiIconKind.Document, () => Run(RemovePackage), false));
            Add(Command("Đặt phiên bản ưu tiên trong kho", EstimateUiIconKind.Check, () => Run(SetPreferred), false));
            Add(Command("Đọc lại kho dữ liệu", EstimateUiIconKind.Refresh, () => Run(RefreshPackages), false));
            Add(Command("Cài bundle từ thư mục", EstimateUiIconKind.Folder, () => Run(ImportBundle), false));
            Add(Command("Cài gói cập nhật có chữ ký", EstimateUiIconKind.Lock, () => Run(ImportSignedUpdate), false));
            Section("3. Ngày hồ sơ khi pin lần đầu");
            preparedDate = DateField("Ngày lập hồ sơ");
            priceDate = DateField("Ngày giá");
            preparedDate.ValueChanged += (s, e) => InvalidatePreview();
            priceDate.ValueChanged += (s, e) => InvalidatePreview();
            Add(Command("Xem tác động đổi gói", EstimateUiIconKind.Link, () => Run(PreviewPackage)));
            impact = new TextBox { ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical,
                Height = 160, BackColor = Color.White };
            Add(impact);
            apply = Command("Xác nhận áp dụng gói", EstimateUiIconKind.Check, () => Run(ApplyPackage));
            apply.Enabled = false;
            Add(apply); Add(Status);
            ClientSizeChanged += (s, e) => pinned.MaximumSize = new Size(Math.Max(250, ClientSize.Width - 50), 0);
            Run(RefreshPackages);
        }

        private DateTimePicker DateField(string label)
        {
            Add(new Label { Text = label, AutoSize = true });
            var field = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy" };
            Add(field); return field;
        }

        private void InvalidatePreview()
        {
            preview = null;
            if (apply != null) apply.Enabled = false;
            if (impact != null) impact.Text = string.Empty;
        }

        private void RefreshPackages()
        {
            InvalidatePreview();
            ProjectProfile profile;
            bool hasProfile = WorkbookProjectProfileService.TryLoad(workbook, out profile);
            if (!hasProfile)
                pinned.Text = "Chưa cấu hình gói pháp lý.";
            else
            {
                var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
                var verify = RegulationPackagePinService.Verify(profile, store);
                pinned.Text = "Trạng thái: " + verify.Status + "\r\n" + profile.RegulationPackageId +
                    " @ " + profile.RegulationPackageVersion + "\r\nSHA-256: " + profile.RegulationPackageChecksum +
                    "\r\n" + verify.Message;
                if (profile.PreparedDate.HasValue) preparedDate.Value = profile.PreparedDate.Value;
                if (profile.PriceDate.HasValue) priceDate.Value = profile.PriceDate.Value;
            }
            preparedDate.Enabled = priceDate.Enabled = !hasProfile;
            string[] issues;
            var available = WorkbookEstimateV2PackageService.ListAvailable(out issues);
            var previous = packages.CurrentRow?.Tag as PackageChoice;
            packages.Rows.Clear();
            foreach (var package in available)
            {
                int index = packages.Rows.Add(package.PackageId, package.DataVersion, package.EffectiveFrom.ToString("dd/MM/yyyy"),
                    package.Modules.FirstOrDefault(m => m.Kind == RegulationModuleKind.Norm)?.RecordCount ?? 0);
                packages.Rows[index].Tag = new PackageChoice(package);
                if (previous != null && previous.Package.PackageChecksum == package.PackageChecksum)
                    packages.CurrentCell = packages.Rows[index].Cells[0];
            }
            // Selecting a row never changes the workbook's pinned package.
            SetStatus(issues.Length == 0 ? available.Count + " gói đã kiểm tra checksum. Online chưa bật." :
                string.Join("\r\n", issues), issues.Length > 0);
            pinned.MaximumSize = new Size(Math.Max(250, ClientSize.Width - 50), 0);
        }

        private void PreviewPackage()
        {
            var selected = packages.CurrentRow?.Tag as PackageChoice;
            if (selected == null) throw new InvalidOperationException("Chọn gói đích đã cài trên máy.");
            preview = WorkbookEstimateV2PackageService.Preview(workbook, selected.Package,
                preparedDate.Value, priceDate.Value);
            impact.Text = "Gói đích: " + selected.Text + "\r\nSHA-256: " + selected.Package.PackageChecksum +
                "\r\nBinding thay đổi: " + preview.Plan.AffectedCount + "\r\n\r\n" +
                string.Join("\r\n", preview.Differences.Concat(preview.Plan.Errors));
            apply.Enabled = preview.Plan.CanApply;
            SetStatus(preview.Plan.CanApply ? "Preview sẵn sàng; workbook chưa thay đổi." :
                "Gói đích thiếu định mức/variant đang dùng. Không thể áp dụng.", !preview.Plan.CanApply);
        }

        private void ApplyPackage()
        {
            if (preview == null) throw new InvalidOperationException("Hãy xem tác động trước.");
            if (MessageBox.Show(this, impact.Text + "\r\n\r\nÁp dụng thay đổi và tạo bản dự phòng?",
                "Đổi gói pháp lý", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            string backup = WorkbookEstimateV2PackageService.Apply(workbook, preview, true);
            RefreshPackages();
            SetStatus("Đã áp dụng gói. Cần cập nhật VL-NC-M, DG và THKP trước khi xuất. Bản dự phòng: " + backup);
        }

        private void ImportBundle()
        {
            using (var dialog = new FolderBrowserDialog { Description = "Chọn thư mục bundle chứa manifest.ttbmanifest" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                new RegulationPackageStore(AppPaths.RegulationPackageDirectory).ImportFromDirectory(dialog.SelectedPath);
                RefreshPackages();
                SetStatus("Đã cài bundle và kiểm tra checksum. Package workbook không thay đổi.");
            }
        }

        private RegulationPackage SelectedPackage()
        {
            var selected = packages.CurrentRow?.Tag as PackageChoice;
            if (selected == null) throw new InvalidOperationException("Chọn một gói trong bảng.");
            return selected.Package;
        }

        private void EditPackage(bool create)
        {
            var selected = SelectedPackage();
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            var bundle = store.LoadBundleRequired(selected.PackageId, selected.DataVersion, selected.PackageChecksum);
            using (var editor = new RegulationPackageEditorForm(bundle, create))
            {
                editor.ShowDialog(this);
                if (editor.InstalledPackage == null) return;
                RefreshPackages();
                foreach (DataGridViewRow row in packages.Rows)
                    if (((PackageChoice)row.Tag).Package.PackageChecksum == editor.InstalledPackage.PackageChecksum)
                        packages.CurrentCell = row.Cells[0];
                SetStatus("Đã cài gói mới. Xem tác động và xác nhận để chọn cho workbook.");
            }
        }

        private void SetPreferred()
        {
            var selected = SelectedPackage();
            if (MessageBox.Show(this, "Đặt " + selected.PackageId + " @ " + selected.DataVersion +
                " làm phiên bản ưu tiên trong kho? Workbook đang pin không thay đổi.", "Gói ưu tiên", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            new RegulationPackageActivationStore(store.RootDirectory).SetPreferred(selected, store.ListInstalled());
            SetStatus("Đã đặt phiên bản ưu tiên. Workbook giữ nguyên gói đang dùng.");
        }

        private void RemovePackage()
        {
            var selected = SelectedPackage();
            var store = new RegulationPackageStore(AppPaths.RegulationPackageDirectory);
            var protectedPackages = new List<RegulationPackage>();
            Excel.Workbooks books = null;
            try
            {
                books = workbook.Application.Workbooks;
                for (int i = 1; i <= books.Count; i++)
                {
                    Excel.Workbook book = null;
                    try
                    {
                        book = books.Item[i]; ProjectProfile profile;
                        if (WorkbookProjectProfileService.TryLoad(book, out profile))
                            protectedPackages.Add(store.LoadRequired(profile.RegulationPackageId, profile.RegulationPackageVersion, profile.RegulationPackageChecksum));
                    }
                    finally { if (book != null) Marshal.ReleaseComObject(book); }
                }
                if (MessageBox.Show(this, "Xóa " + selected.PackageId + " @ " + selected.DataVersion +
                    " khỏi kho?\nGói được lưu dự phòng. Hồ sơ đang đóng hoặc mở ở phiên Excel khác có thể dùng gói này; hãy kiểm tra trước.\nGói đang pin trong phiên Excel này hoặc đang ưu tiên sẽ bị chặn.",
                    "Xóa gói pháp lý", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                string archive = store.Remove(selected, protectedPackages);
                RefreshPackages(); SetStatus("Đã xóa khỏi kho. Có thể nhập lại từ bản dự phòng: " + archive);
            }
            finally { if (books != null) Marshal.ReleaseComObject(books); }
        }

        private void ImportSignedUpdate()
        {
            using (var dialog = new OpenFileDialog { Filter = "Gói cập nhật (*.ttbupdate)|*.ttbupdate" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                Version version;
                if (!Version.TryParse(AppInfo.Version, out version)) version = new Version(1, 0, 0, 0);
                var service = new OfflineUpdateCenterService(AppPaths.RegulationPackageDirectory,
                    OfflineUpdateTrustCatalog.Production, version);
                ProjectProfile profile;
                RegulationPackage current = null;
                if (WorkbookProjectProfileService.TryLoad(workbook, out profile))
                    current = RegulationPackagePinService.Verify(profile,
                        new RegulationPackageStore(AppPaths.RegulationPackageDirectory)).Package;
                var inspection = service.Inspect(dialog.FileName, current);
                if (MessageBox.Show(this, "Cài gói " + inspection.Verification.Package.PackageId +
                    " @ " + inspection.Verification.Package.DataVersion + "?\r\nPackage workbook sẽ giữ nguyên.",
                    "Cài dữ liệu offline", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                service.InstallAndActivate(inspection);
                RefreshPackages();
                SetStatus("Đã cài gói có chữ ký. Workbook giữ nguyên package đang pin.");
            }
        }

        private sealed class PackageChoice
        {
            internal PackageChoice(RegulationPackage package) { Package = package; }
            internal RegulationPackage Package { get; }
            public string Text => Package.PackageId + " @ " + Package.DataVersion;
        }
    }
}
