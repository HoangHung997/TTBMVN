using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class PackageMigrationControl : UserControl
    {
        private readonly Excel.Workbook workbook;
        private readonly IReadOnlyList<RegulationPackage> packages;
        private readonly ComboBox targetCombo;
        private readonly Label sourceLabel;
        private readonly Label statusLabel;
        private readonly Button applyButton;
        private readonly DataGridView valueGrid;
        private readonly DataGridView packageGrid;

        public PackageMigrationControl(Excel.Workbook workbook)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            packages = RegulationPackageBootstrapService.LoadAvailablePackages();
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            Padding = new Padding(14);

            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 82,
                ColumnCount = 5,
                RowCount = 2
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 52));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 196));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            header.Controls.Add(Label("Đang dùng"), 0, 0);
            sourceLabel = Label(string.Empty);
            sourceLabel.Font = new Font(Font, FontStyle.Bold);
            header.Controls.Add(sourceLabel, 1, 0);
            header.Controls.Add(Label("Gói đích"), 2, 0);
            targetCombo = new ComboBox
            {
                Name = "cboMigrationTarget",
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                DisplayMember = nameof(PackageItem.DisplayName)
            };
            foreach (RegulationPackage package in packages
                .OrderBy(item => item.EffectiveFrom)
                .ThenBy(item => item.DataVersion, StringComparer.Ordinal))
            {
                targetCombo.Items.Add(new PackageItem(package));
            }
            header.Controls.Add(targetCombo, 3, 0);

            var previewButton = new Button
            {
                Name = "btnPreviewMigration",
                Text = "Xem trước",
                Width = 86,
                Height = 30
            };
            applyButton = new Button
            {
                Name = "btnApplyMigration",
                Text = "Chuyển gói",
                Width = 94,
                Height = 30,
                Enabled = false
            };
            var commands = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            commands.Controls.Add(previewButton);
            commands.Controls.Add(applyButton);
            header.Controls.Add(commands, 4, 0);
            header.SetRowSpan(commands, 2);

            statusLabel = new Label
            {
                Name = "lblMigrationStatus",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(55, 78, 61)
            };
            header.Controls.Add(statusLabel, 0, 1);
            header.SetColumnSpan(statusLabel, 4);

            valueGrid = Grid("dgvMigrationValues");
            valueGrid.Columns.Add(Column("item", "Kết quả", 210));
            valueGrid.Columns.Add(Column("source", "Trước chuyển", 145));
            valueGrid.Columns.Add(Column("target", "Sau chuyển", 145));
            DataGridViewTextBoxColumn difference = Column("difference", "Chênh lệch", 145);
            difference.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            valueGrid.Columns.Add(difference);

            packageGrid = Grid("dgvMigrationPackageChanges");
            packageGrid.Columns.Add(Column("impact", "Ảnh hưởng", 110));
            packageGrid.Columns.Add(Column("kind", "Loại", 142));
            packageGrid.Columns.Add(Column("key", "Mục", 155));
            packageGrid.Columns.Add(Column("old", "Gói nguồn", 190));
            DataGridViewTextBoxColumn target = Column("new", "Gói đích", 190);
            target.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            packageGrid.Columns.Add(target);

            var tabs = new TabControl { Dock = DockStyle.Fill };
            var resultPage = new TabPage("Kết quả tính") { Padding = new Padding(4) };
            var packagePage = new TabPage("Thay đổi package") { Padding = new Padding(4) };
            resultPage.Controls.Add(valueGrid);
            packagePage.Controls.Add(packageGrid);
            tabs.TabPages.Add(resultPage);
            tabs.TabPages.Add(packagePage);

            Controls.Add(tabs);
            Controls.Add(header);
            previewButton.Click += (sender, args) => PreviewSelected();
            applyButton.Click += (sender, args) => ApplyFromDialog();
            targetCombo.SelectedIndexChanged += (sender, args) =>
            {
                CurrentPreview = null;
                applyButton.Enabled = false;
                statusLabel.Text = "Bấm Xem trước để tính ảnh hưởng trước khi chuyển gói.";
            };
            LoadCurrentPackage();
        }

        public WorkbookPackageMigrationPreview CurrentPreview { get; private set; }

        public bool SelectTarget(string packageId, string dataVersion)
        {
            for (int index = 0; index < targetCombo.Items.Count; index++)
            {
                PackageItem item = targetCombo.Items[index] as PackageItem;
                if (item != null &&
                    string.Equals(item.Package.PackageId, packageId, StringComparison.Ordinal) &&
                    string.Equals(item.Package.DataVersion, dataVersion, StringComparison.Ordinal))
                {
                    targetCombo.SelectedIndex = index;
                    return true;
                }
            }
            return false;
        }

        public WorkbookPackageMigrationPreview PreviewSelected()
        {
            PackageItem item = targetCombo.SelectedItem as PackageItem;
            if (item == null)
                return null;
            try
            {
                Cursor = Cursors.WaitCursor;
                CurrentPreview = WorkbookPackageMigrationService.Preview(
                    workbook,
                    AppPaths.RegulationPackageDirectory,
                    item.Package.PackageId,
                    item.Package.DataVersion,
                    item.Package.PackageChecksum);
                BindPreview(CurrentPreview);
                return CurrentPreview;
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, "Preview package migration UI", "DT-504", "preview", workbook, string.Empty);
                CurrentPreview = null;
                applyButton.Enabled = false;
                statusLabel.ForeColor = Color.FromArgb(165, 42, 42);
                statusLabel.Text = ex.Message;
                return null;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public WorkbookPackageMigrationResult ApplyPreview(string backupPath, bool confirmed)
        {
            if (CurrentPreview == null)
                throw new InvalidOperationException("Chua co preview migration.");
            WorkbookPackageMigrationResult result = WorkbookPackageMigrationService.Apply(
                workbook,
                AppPaths.RegulationPackageDirectory,
                CurrentPreview,
                backupPath,
                confirmed);
            if (result.Status == WorkbookPackageMigrationStatus.Applied)
            {
                LoadCurrentPackage();
                CurrentPreview = null;
                applyButton.Enabled = false;
                statusLabel.ForeColor = Color.FromArgb(35, 92, 55);
                statusLabel.Text = "Đã chuyển package và lưu backup: " + result.BackupPath;
            }
            return result;
        }

        private void ApplyFromDialog()
        {
            if (CurrentPreview == null || !CurrentPreview.Impact.CanApply)
                return;
            string workbookDirectory = string.IsNullOrWhiteSpace(workbook.Path)
                ? AppDomain.CurrentDomain.BaseDirectory
                : workbook.Path;
            string extension = Path.GetExtension(workbook.Name);
            string baseName = Path.GetFileNameWithoutExtension(workbook.Name);
            using (var dialog = new SaveFileDialog
            {
                Title = "Lưu bản sao trước khi chuyển package",
                Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                InitialDirectory = workbookDirectory,
                FileName = baseName + ".before-migration-" +
                    DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + extension,
                AddExtension = true,
                OverwritePrompt = true
            })
            {
                if (dialog.ShowDialog(FindForm()) != DialogResult.OK)
                    return;
                DialogResult confirmation = MessageBox.Show(
                    FindForm(),
                    "Ứng dụng sẽ ghi lại phụ lục, tổng hợp kinh phí và audit theo package đích. Tiếp tục?",
                    "Xác nhận chuyển package",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (confirmation != DialogResult.Yes)
                    return;
                try
                {
                    Cursor = Cursors.WaitCursor;
                    ApplyPreview(dialog.FileName, true);
                }
                catch (Exception ex)
                {
                    RuntimeLogger.LogOperation(
                        ex, "Apply package migration UI", "DT-504", "apply", workbook, string.Empty);
                    statusLabel.ForeColor = Color.FromArgb(165, 42, 42);
                    statusLabel.Text = ex.Message;
                    MessageBox.Show(
                        FindForm(),
                        ex.Message,
                        "Không chuyển được package",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void BindPreview(WorkbookPackageMigrationPreview preview)
        {
            valueGrid.Rows.Clear();
            foreach (WorkbookPackageMigrationValueChange change in preview.Impact.ValueChanges)
            {
                valueGrid.Rows.Add(
                    change.DisplayName,
                    FormatNumber(change.SourceValue),
                    FormatNumber(change.TargetValue),
                    FormatSigned(change.Difference));
            }
            packageGrid.Rows.Clear();
            foreach (RegulationPackageChange change in preview.Plan.Diff.Changes)
            {
                packageGrid.Rows.Add(
                    ImpactText(change.Impact),
                    change.Kind,
                    change.Key,
                    change.OldValue,
                    change.NewValue);
            }
            applyButton.Enabled = preview.Impact.CanApply && preview.Plan.Diff.HasChanges;
            statusLabel.ForeColor = preview.Impact.CanApply
                ? Color.FromArgb(35, 92, 55)
                : Color.FromArgb(165, 42, 42);
            statusLabel.Text = preview.Impact.CanApply
                ? "Sẵn sàng: " + preview.Impact.EstimateLineCount + " dòng, " +
                    preview.Impact.AuditEntryCount + " audit entry, " +
                    preview.Plan.Diff.Changes.Count + " thay đổi package."
                : string.Join(" ", preview.Impact.Issues.Take(2));
        }

        private void LoadCurrentPackage()
        {
            ProjectProfile profile = WorkbookProjectProfileService.LoadRequired(workbook);
            sourceLabel.Text = profile.RegulationPackageId + " @ " + profile.RegulationPackageVersion;
            PackageItem preferred = targetCombo.Items.Cast<object>()
                .OfType<PackageItem>()
                .Where(item => !string.Equals(
                    item.Package.PackageChecksum,
                    profile.RegulationPackageChecksum,
                    StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(item => item.Package.EffectiveFrom)
                .FirstOrDefault();
            targetCombo.SelectedItem = preferred ?? targetCombo.Items.Cast<object>().FirstOrDefault();
        }

        private string FormatNumber(decimal value)
        {
            return value.ToString("N0", ExcelCulture.GetNumberCulture(workbook.Application));
        }

        private string FormatSigned(decimal value)
        {
            string prefix = value > 0m ? "+" : string.Empty;
            return prefix + FormatNumber(value);
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

        private static Label Label(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
        }

        private static DataGridView Grid(string name)
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

        private static DataGridViewTextBoxColumn Column(string name, string text, int width)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = text,
                Width = width,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private sealed class PackageItem
        {
            internal PackageItem(RegulationPackage package)
            {
                Package = package;
                DisplayName = package.PackageId + " @ " + package.DataVersion;
            }

            internal RegulationPackage Package { get; }
            public string DisplayName { get; }
        }
    }
}
