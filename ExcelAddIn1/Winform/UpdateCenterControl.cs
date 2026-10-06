using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class UpdateCenterControl : UserControl
    {
        private readonly Excel.Workbook workbook;
        private readonly OfflineUpdateCenterService service;
        private readonly RegulationPackage currentWorkbookPackage;
        private readonly Label currentLabel;
        private readonly Label updateLabel;
        private readonly Label statusLabel;
        private readonly Button installButton;
        private readonly Button activateButton;
        private readonly Button migrateButton;
        private readonly DataGridView installedGrid;
        private readonly DataGridView diffGrid;

        public UpdateCenterControl(Excel.Workbook workbook)
            : this(
                workbook,
                CreateDefaultService(),
                LoadCurrentPackage(workbook))
        {
        }

        public UpdateCenterControl(
            Excel.Workbook workbook,
            OfflineUpdateCenterService service,
            RegulationPackage currentWorkbookPackage)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.service = service ?? throw new ArgumentNullException(nameof(service));
            this.currentWorkbookPackage = currentWorkbookPackage ??
                throw new ArgumentNullException(nameof(currentWorkbookPackage));

            Dock = DockStyle.Fill;
            BackColor = Color.White;
            Padding = new Padding(14);

            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 118,
                ColumnCount = 4,
                RowCount = 3
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 106));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));

            header.Controls.Add(CreateLabel("Workbook"), 0, 0);
            currentLabel = CreateLabel(
                currentWorkbookPackage.PackageId + " @ " + currentWorkbookPackage.DataVersion);
            currentLabel.Font = new Font(Font, FontStyle.Bold);
            header.Controls.Add(currentLabel, 1, 0);
            var onlineLabel = CreateLabel("Online: Chưa bật");
            onlineLabel.ForeColor = Color.DimGray;
            header.Controls.Add(onlineLabel, 2, 0);
            var onlineButton = new Button
            {
                Name = "btnCheckOnlineUpdate",
                Text = "Kiểm tra online",
                Dock = DockStyle.Fill,
                Enabled = false
            };
            header.Controls.Add(onlineButton, 3, 0);

            header.Controls.Add(CreateLabel("Gói offline"), 0, 1);
            updateLabel = CreateLabel("Chưa chọn gói cập nhật");
            updateLabel.AutoEllipsis = true;
            header.Controls.Add(updateLabel, 1, 1);
            var browseButton = new Button
            {
                Name = "btnBrowseOfflineUpdate",
                Text = "Chọn gói",
                Dock = DockStyle.Fill
            };
            header.Controls.Add(browseButton, 2, 1);
            installButton = new Button
            {
                Name = "btnInstallOfflineUpdate",
                Text = "Cài và dùng",
                Dock = DockStyle.Fill,
                Enabled = false
            };
            header.Controls.Add(installButton, 3, 1);

            statusLabel = CreateLabel("Sẵn sàng cập nhật offline.");
            statusLabel.Name = "lblUpdateCenterStatus";
            statusLabel.ForeColor = Color.FromArgb(55, 78, 61);
            header.Controls.Add(statusLabel, 0, 2);
            header.SetColumnSpan(statusLabel, 4);

            installedGrid = CreateGrid("dgvInstalledUpdates");
            installedGrid.Columns.Add(CreateColumn("preferred", "Mặc định", 78));
            installedGrid.Columns.Add(CreateColumn("package", "Package", 180));
            installedGrid.Columns.Add(CreateColumn("version", "Version", 82));
            installedGrid.Columns.Add(CreateColumn("effective", "Hiệu lực", 94));
            installedGrid.Columns.Add(CreateColumn("workbook", "Workbook", 88));
            DataGridViewTextBoxColumn checksum = CreateColumn("checksum", "Checksum", 230);
            checksum.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            installedGrid.Columns.Add(checksum);

            diffGrid = CreateGrid("dgvOfflineUpdateDiff");
            diffGrid.Columns.Add(CreateColumn("impact", "Ảnh hưởng", 112));
            diffGrid.Columns.Add(CreateColumn("kind", "Loại", 138));
            diffGrid.Columns.Add(CreateColumn("key", "Mục", 150));
            diffGrid.Columns.Add(CreateColumn("old", "Hiện tại", 185));
            DataGridViewTextBoxColumn next = CreateColumn("new", "Gói cập nhật", 185);
            next.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            diffGrid.Columns.Add(next);

            var installedPage = new TabPage("Gói đã cài") { Padding = new Padding(4) };
            installedPage.Controls.Add(installedGrid);
            var diffPage = new TabPage("Thay đổi") { Padding = new Padding(4) };
            diffPage.Controls.Add(diffGrid);
            var tabs = new TabControl { Dock = DockStyle.Fill };
            tabs.TabPages.Add(installedPage);
            tabs.TabPages.Add(diffPage);

            var commands = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0)
            };
            migrateButton = new Button
            {
                Name = "btnMigrateToInstalledPackage",
                Text = "Chuyển workbook",
                Width = 128,
                Height = 30,
                Enabled = false
            };
            activateButton = new Button
            {
                Name = "btnActivateInstalledPackage",
                Text = "Đặt mặc định",
                Width = 112,
                Height = 30,
                Enabled = false
            };
            var refreshButton = new Button
            {
                Name = "btnRefreshInstalledPackages",
                Text = "Làm mới",
                Width = 88,
                Height = 30
            };
            commands.Controls.Add(migrateButton);
            commands.Controls.Add(activateButton);
            commands.Controls.Add(refreshButton);

            Controls.Add(tabs);
            Controls.Add(commands);
            Controls.Add(header);

            browseButton.Click += (sender, args) => BrowseAndInspect();
            installButton.Click += (sender, args) => InstallCurrent();
            refreshButton.Click += (sender, args) => RefreshInstalled();
            activateButton.Click += (sender, args) => ActivateSelected();
            migrateButton.Click += (sender, args) => RequestMigration();
            installedGrid.SelectionChanged += (sender, args) => UpdateSelectionCommands();
            RefreshInstalled();
        }

        public event Action<RegulationPackage> MigrationRequested;

        public OfflineUpdateInspection CurrentInspection { get; private set; }

        public OfflineUpdateInspection InspectFile(string path)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                CurrentInspection = service.Inspect(path, currentWorkbookPackage);
                BindInspection(CurrentInspection);
                SaveLastDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
                return CurrentInspection;
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Inspect offline update", "DT-602", "inspect", workbook, string.Empty);
                CurrentInspection = null;
                installButton.Enabled = false;
                diffGrid.Rows.Clear();
                SetStatus(ex.Message, false);
                return null;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public OfflineUpdateInstallResult InstallCurrent()
        {
            if (CurrentInspection == null)
                return null;
            try
            {
                Cursor = Cursors.WaitCursor;
                OfflineUpdateInstallResult result = service.InstallAndActivate(CurrentInspection);
                RefreshInstalled();
                SetStatus(
                    result.Status == OfflineUpdateInstallStatus.InstalledAndActivated
                        ? "Đã cài và đặt package làm mặc định cho dự án mới."
                        : "Package đã có; đã đặt lại làm mặc định cho dự án mới.",
                    true);
                CurrentInspection = service.Inspect(
                    CurrentInspection.Verification.ArchivePath,
                    currentWorkbookPackage);
                BindInspection(CurrentInspection);
                return result;
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Install offline update", "DT-602", "install", workbook, string.Empty);
                SetStatus(ex.Message, false);
                MessageBox.Show(
                    FindForm(),
                    ex.Message,
                    "Không cài được gói cập nhật",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void RefreshInstalled()
        {
            RegulationPackage selected = SelectedPackage;
            installedGrid.Rows.Clear();
            foreach (RegulationPackage package in service.ListInstalled()
                .OrderByDescending(item => item.EffectiveFrom)
                .ThenByDescending(item => item.DataVersion, StringComparer.Ordinal))
            {
                bool preferred = service.IsPreferred(package);
                bool workbookPackage = string.Equals(
                    package.PackageChecksum,
                    currentWorkbookPackage.PackageChecksum,
                    StringComparison.OrdinalIgnoreCase);
                int rowIndex = installedGrid.Rows.Add(
                    preferred ? "Có" : string.Empty,
                    package.PackageId,
                    package.DataVersion,
                    package.EffectiveFrom.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
                    workbookPackage ? "Đang dùng" : string.Empty,
                    package.PackageChecksum);
                installedGrid.Rows[rowIndex].Tag = package;
                if (selected != null && SameIdentity(selected, package))
                    installedGrid.Rows[rowIndex].Selected = true;
            }
            if (installedGrid.SelectedRows.Count == 0 && installedGrid.Rows.Count > 0)
                installedGrid.Rows[0].Selected = true;
            UpdateSelectionCommands();
        }

        private RegulationPackage SelectedPackage
        {
            get
            {
                return installedGrid.SelectedRows.Count == 0
                    ? null
                    : installedGrid.SelectedRows[0].Tag as RegulationPackage;
            }
        }

        private void BrowseAndInspect()
        {
            string initialDirectory = LoadLastDirectory();
            using (var dialog = new OpenFileDialog
            {
                Title = "Chọn gói cập nhật offline",
                Filter = "TTBMVN update (*.ttbupdate)|*.ttbupdate|All files (*.*)|*.*",
                InitialDirectory = initialDirectory,
                CheckFileExists = true,
                Multiselect = false
            })
            {
                if (dialog.ShowDialog(FindForm()) == DialogResult.OK)
                    InspectFile(dialog.FileName);
            }
        }

        private void BindInspection(OfflineUpdateInspection inspection)
        {
            OfflineUpdateVerificationResult verification = inspection.Verification;
            updateLabel.Text = verification.Package.PackageId + " @ " +
                verification.Package.DataVersion + " | " + verification.Manifest.SigningKeyId;
            diffGrid.Rows.Clear();
            if (inspection.Diff != null)
            {
                foreach (RegulationPackageChange change in inspection.Diff.Changes)
                {
                    diffGrid.Rows.Add(
                        ImpactText(change.Impact),
                        change.Kind,
                        change.Key,
                        change.OldValue,
                        change.NewValue);
                }
            }
            installButton.Enabled = verification.Disposition == OfflineUpdateDisposition.Ready ||
                verification.Disposition == OfflineUpdateDisposition.AlreadyInstalled;
            installButton.Text = verification.Disposition == OfflineUpdateDisposition.AlreadyInstalled
                ? "Dùng bản này"
                : "Cài và dùng";
            SetStatus(verification.DispositionReason, installButton.Enabled);
        }

        private void ActivateSelected()
        {
            RegulationPackage package = SelectedPackage;
            if (package == null || service.IsPreferred(package))
                return;
            try
            {
                service.ActivateInstalled(package);
                RefreshInstalled();
                SetStatus(
                    "Đã rollback bản mặc định về " + package.PackageId + " @ " +
                    package.DataVersion + ". Workbook hiện tại không thay đổi.",
                    true);
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Activate installed package", "DT-602", "rollback", workbook, string.Empty);
                SetStatus(ex.Message, false);
            }
        }

        private void RequestMigration()
        {
            RegulationPackage package = SelectedPackage;
            if (package != null)
                MigrationRequested?.Invoke(package);
        }

        private void UpdateSelectionCommands()
        {
            RegulationPackage package = SelectedPackage;
            activateButton.Enabled = package != null && !service.IsPreferred(package);
            migrateButton.Enabled = package != null && !string.Equals(
                package.PackageChecksum,
                currentWorkbookPackage.PackageChecksum,
                StringComparison.OrdinalIgnoreCase);
        }

        private void SetStatus(string text, bool success)
        {
            statusLabel.Text = text;
            statusLabel.ForeColor = success
                ? Color.FromArgb(35, 92, 55)
                : Color.FromArgb(165, 42, 42);
        }

        private static string LoadLastDirectory()
        {
            try
            {
                if (File.Exists(AppPaths.UpdateCenterStatePath))
                {
                    string encoded = File.ReadAllText(AppPaths.UpdateCenterStatePath, Encoding.UTF8).Trim();
                    string path = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
                    if (Directory.Exists(path))
                        return path;
                }
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Load update center state");
            }
            Directory.CreateDirectory(AppPaths.UpdateDirectory);
            return AppPaths.UpdateDirectory;
        }

        private static void SaveLastDirectory(string directory)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
                    return;
                AppPaths.EnsureRoot();
                File.WriteAllText(
                    AppPaths.UpdateCenterStatePath,
                    Convert.ToBase64String(Encoding.UTF8.GetBytes(Path.GetFullPath(directory))),
                    new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Save update center state");
            }
        }

        private static Version ParseAppVersion(string value)
        {
            Version result;
            return Version.TryParse(value, out result) ? result : new Version(0, 0, 0, 0);
        }

        private static OfflineUpdateCenterService CreateDefaultService()
        {
            RegulationPackageBootstrapService.LoadAvailablePackages();
            return new OfflineUpdateCenterService(
                AppPaths.RegulationPackageDirectory,
                OfflineUpdateTrustCatalog.Production,
                ParseAppVersion(AppInfo.Version));
        }

        private static RegulationPackage LoadCurrentPackage(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            return WorkbookRegulationPackageService.LoadPinnedPackageRequired(
                workbook,
                AppPaths.RegulationPackageDirectory);
        }

        private static string ImpactText(RegulationPackageChangeImpact impact)
        {
            switch (impact)
            {
                case RegulationPackageChangeImpact.CalculationData: return "Dữ liệu tính";
                case RegulationPackageChangeImpact.LegalBasis: return "Căn cứ";
                default: return "Thông tin";
            }
        }

        private static bool SameIdentity(RegulationPackage left, RegulationPackage right)
        {
            return left != null && right != null &&
                string.Equals(left.PackageId, right.PackageId, StringComparison.Ordinal) &&
                string.Equals(left.DataVersion, right.DataVersion, StringComparison.Ordinal) &&
                string.Equals(left.PackageChecksum, right.PackageChecksum, StringComparison.OrdinalIgnoreCase);
        }

        private static Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
        }

        private static DataGridView CreateGrid(string name)
        {
            return new DataGridView
            {
                Name = name,
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                MultiSelect = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
        }

        private static DataGridViewTextBoxColumn CreateColumn(string name, string text, int width)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = text,
                Width = width,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }
    }
}
