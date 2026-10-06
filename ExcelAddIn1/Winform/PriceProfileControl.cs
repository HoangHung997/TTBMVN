using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class PriceProfileControl : UserControl
    {
        private readonly Excel.Workbook workbook;
        private readonly CultureInfo numberCulture;
        private readonly TextBox profileIdInput;
        private readonly TextBox versionInput;
        private readonly TextBox displayNameInput;
        private readonly TextBox locationInput;
        private readonly DateTimePicker valuationDateInput;
        private readonly ComboBox audienceInput;
        private readonly DataGridView priceGrid;
        private readonly Label statusLabel;
        private readonly Label checksumLabel;
        private readonly Dictionary<MachineRateAudience, PriceProfile> profileDrafts =
            new Dictionary<MachineRateAudience, PriceProfile>();
        private PriceProfile currentProfile;
        private MachineRateAudience activeAudience;
        private bool loading;
        private bool dirty;

        public PriceProfileControl(Excel.Workbook workbook)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            numberCulture = ExcelCulture.GetNumberCulture(workbook.Application);
            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;

            var metadata = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 118,
                Padding = new Padding(10, 7, 10, 3),
                ColumnCount = 6,
                RowCount = 4
            };
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10));
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            metadata.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            metadata.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            metadata.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            metadata.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            profileIdInput = CreateTextBox("txtPriceProfileId");
            versionInput = CreateTextBox("txtPriceProfileVersion");
            displayNameInput = CreateTextBox("txtPriceProfileName");
            locationInput = CreateTextBox("txtPriceLocation");
            valuationDateInput = new DateTimePicker
            {
                Name = "dtpPriceDate",
                Dock = DockStyle.Fill,
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "dd/MM/yyyy"
            };
            audienceInput = new ComboBox
            {
                Name = "cboPriceAudience",
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            audienceInput.Items.Add(new OptionItem<MachineRateAudience>(
                "Hưởng lương NSNN", MachineRateAudience.StateBudgetSalary));
            audienceInput.Items.Add(new OptionItem<MachineRateAudience>(
                "Không hưởng lương NSNN", MachineRateAudience.NonStateSalary));

            AddMetadata(metadata, 0, 0, "Mã hồ sơ giá", profileIdInput);
            AddMetadata(metadata, 1, 0, "Phiên bản", versionInput);
            AddMetadata(metadata, 2, 0, "Tên hồ sơ", displayNameInput);
            AddMetadata(metadata, 3, 0, "Địa điểm", locationInput);
            AddMetadata(metadata, 4, 0, "Ngày giá", valuationDateInput);
            AddMetadata(metadata, 5, 0, "Đối tượng lương", audienceInput);

            checksumLabel = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                ForeColor = SystemColors.GrayText,
                TextAlign = ContentAlignment.MiddleLeft
            };
            metadata.Controls.Add(new Label
            {
                Text = "Checksum snapshot",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, 2);
            metadata.Controls.Add(checksumLabel, 0, 3);
            metadata.SetColumnSpan(checksumLabel, 6);

            priceGrid = CreateGrid();
            var gridPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 4, 10, 4) };
            gridPanel.Controls.Add(priceGrid);

            var commands = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                Padding = new Padding(8, 4, 8, 3),
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };
            Button saveButton = CreateButton("btnSavePriceProfile", "Lưu hồ sơ", 100);
            Button exportButton = CreateButton("btnExportPriceProfile", "Export", 82);
            Button importButton = CreateButton("btnImportPriceProfile", "Import", 82);
            Button reloadButton = CreateButton("btnReloadLegacyPrices", "Đọc VL-NC-M", 112);
            Button checkButton = CreateButton("btnCheckPriceCoverage", "Kiểm tra", 90);
            Button addButton = CreateButton("btnAddPrice", "Thêm", 72);
            Button deleteButton = CreateButton("btnDeletePrice", "Xóa", 72);
            saveButton.Click += (sender, args) => Execute("Lưu hồ sơ giá", () => SaveCurrentProfile(true));
            exportButton.Click += (sender, args) => Execute("Export hồ sơ giá", ExportProfile);
            importButton.Click += (sender, args) => Execute("Import hồ sơ giá", ImportProfile);
            reloadButton.Click += (sender, args) => Execute("Đọc VL-NC-M", ReloadLegacyProfile);
            checkButton.Click += (sender, args) => Execute("Kiểm tra bảng giá", ShowCoverage);
            addButton.Click += (sender, args) => AddRow();
            deleteButton.Click += (sender, args) => DeleteSelectedRows();
            commands.Controls.Add(saveButton);
            commands.Controls.Add(exportButton);
            commands.Controls.Add(importButton);
            commands.Controls.Add(reloadButton);
            commands.Controls.Add(checkButton);
            commands.Controls.Add(deleteButton);
            commands.Controls.Add(addButton);

            statusLabel = new Label
            {
                Name = "lblPriceStatus",
                Dock = DockStyle.Bottom,
                Height = 30,
                Padding = new Padding(10, 5, 10, 3),
                AutoEllipsis = true,
                ForeColor = SystemColors.GrayText
            };

            Controls.Add(gridPanel);
            Controls.Add(metadata);
            Controls.Add(statusLabel);
            Controls.Add(commands);

            HookDirtyEvents();
            audienceInput.SelectionChangeCommitted += (sender, args) => SwitchAudience();
            LoadInitialProfile();
        }

        public PriceProfile CurrentProfile => currentProfile;

        public PriceProfile BuildCurrentProfile()
        {
            return BuildCurrentProfile(SelectedAudience);
        }

        private PriceProfile BuildCurrentProfile(MachineRateAudience audience)
        {
            priceGrid.EndEdit();
            string profileId = profileIdInput.Text.Trim();
            string version = versionInput.Text.Trim();
            var entries = new List<PriceProfileEntry>();
            var overrides = new List<PriceProfileOverride>();
            DateTime createdAt = currentProfile != null && !dirty &&
                string.Equals(currentProfile.ProfileId, profileId, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(currentProfile.DataVersion, version, StringComparison.Ordinal)
                ? currentProfile.CreatedAtUtc
                : DateTime.UtcNow;

            foreach (DataGridViewRow row in priceGrid.Rows)
            {
                if (row.IsNewRow)
                    continue;
                string code = CellText(row, "colPriceCode");
                string name = CellText(row, "colPriceName");
                if (code.Length == 0 && name.Length == 0)
                    continue;
                PriceResourceKind kind = Value<PriceResourceKind>(row, "colPriceKind");
                string unit = CellText(row, "colPriceUnit");
                decimal basePrice = Money(row, "colBasePrice");
                DateTime sourceDate = Date(row, "colSourceDate");
                PriceSourceKind sourceKind = Value<PriceSourceKind>(row, "colSourceKind");
                string source = CellText(row, "colPriceSource");
                string[] aliases = CellText(row, "colAliases")
                    .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(value => value.Trim())
                    .Where(value => value.Length > 0)
                    .ToArray();
                string legacyName = CellText(row, "colLegacyName");
                entries.Add(new PriceProfileEntry(
                    code,
                    kind,
                    name,
                    unit,
                    basePrice,
                    sourceDate,
                    sourceKind,
                    source,
                    aliases,
                    legacyName));
                if (Boolean(row, "colHasOverride"))
                {
                    decimal applied = Money(row, "colAppliedPrice");
                    PriceProfileOverride existing = row.Tag as PriceProfileOverride;
                    bool unchanged = existing != null &&
                        existing.OldUnitPriceVnd == basePrice &&
                        existing.NewUnitPriceVnd == applied &&
                        string.Equals(existing.Reason, CellText(row, "colOverrideReason"), StringComparison.Ordinal) &&
                        string.Equals(existing.SourceReference, CellText(row, "colOverrideSource"), StringComparison.Ordinal);
                    overrides.Add(new PriceProfileOverride(
                        code,
                        basePrice,
                        applied,
                        CellText(row, "colOverrideReason"),
                        CellText(row, "colOverrideSource"),
                        unchanged ? existing.ModifiedBy : Environment.UserName,
                        unchanged ? existing.ModifiedAtUtc : createdAt));
                }
            }

            return PriceProfile.Create(
                profileId,
                version,
                displayNameInput.Text,
                locationInput.Text,
                valuationDateInput.Value.Date,
                audience,
                createdAt,
                entries,
                overrides);
        }

        public void SaveCurrentProfile(bool applyDirectMaterials)
        {
            PriceProfile profile = BuildCurrentProfile();
            var store = new PriceProfileStore(AppPaths.PriceProfileDirectory);
            store.Import(profile);
            profileDrafts[profile.LaborAudience] = profile;
            foreach (PriceProfile draft in profileDrafts.Values)
                store.Import(draft);
            WorkbookPriceProfilePortfolioService.Save(
                workbook,
                PriceProfilePortfolio.Create(profileDrafts.Values));
            WorkbookPriceProfileService.SaveAndPin(workbook, profile);
            LegacyPriceApplyResult apply = applyDirectMaterials &&
                profile.LaborAudience == MachineRateAudience.NonStateSalary
                ? LegacyWorkbookPriceProfileService.ApplyDirectMaterialPrices(workbook, profile)
                : new LegacyPriceApplyResult(0, Array.Empty<string>());
            LoadProfile(profile, false);
            statusLabel.ForeColor = SystemColors.GrayText;
            statusLabel.Text = "Đã lưu hồ sơ " + AudienceText(profile.LaborAudience) + " với " +
                profile.Entries.Count + " giá; cập nhật " + apply.Updated +
                " giá vật liệu vào bảng legacy.";
        }

        private void LoadInitialProfile()
        {
            PriceProfilePortfolio portfolio;
            if (WorkbookPriceProfilePortfolioService.TryLoad(workbook, out portfolio))
            {
                foreach (PriceProfile item in portfolio.Profiles)
                    profileDrafts[item.LaborAudience] = item;
                PriceProfile selected;
                if (!profileDrafts.TryGetValue(MachineRateAudience.NonStateSalary, out selected))
                    selected = profileDrafts.Values.OrderBy(item => item.LaborAudience).First();
                LoadProfile(selected, false);
                statusLabel.Text = "Đã nạp " + profileDrafts.Count +
                    "/2 hồ sơ giá HLNS/KHLNS từ workbook.";
                return;
            }
            PriceProfile profile;
            if (WorkbookPriceProfileService.TryLoad(workbook, out profile))
            {
                profileDrafts[profile.LaborAudience] = profile;
                LoadProfile(profile, false);
                statusLabel.Text = "Đã nạp PriceProfile snapshot từ workbook.";
                return;
            }
            profile = LegacyWorkbookPriceProfileService.Import(workbook);
            profileDrafts[profile.LaborAudience] = profile;
            LoadProfile(profile, false);
            statusLabel.Text = "Đã đọc dữ liệu cũ từ VL-NC-M cho KHLNS; hãy lưu và tạo thêm hồ sơ HLNS khi cần.";
        }

        private void LoadProfile(PriceProfile profile, bool markDirty)
        {
            loading = true;
            try
            {
                currentProfile = profile ?? throw new ArgumentNullException(nameof(profile));
                activeAudience = profile.LaborAudience;
                profileDrafts[activeAudience] = profile;
                profileIdInput.Text = profile.ProfileId;
                versionInput.Text = profile.DataVersion;
                displayNameInput.Text = profile.DisplayName;
                locationInput.Text = profile.Location;
                valuationDateInput.Value = profile.ValuationDate;
                SelectAudience(profile.LaborAudience);
                checksumLabel.Text = profile.Checksum;
                priceGrid.Rows.Clear();
                foreach (PriceProfileEntry entry in profile.Entries)
                {
                    PriceProfilePrice price = profile.FindRequired(entry.Code);
                    int index = priceGrid.Rows.Add(
                        entry.Code,
                        entry.Kind,
                        entry.DisplayName,
                        entry.Unit,
                        entry.BaseUnitPriceVnd,
                        price.IsOverridden,
                        price.AppliedUnitPriceVnd,
                        price.Override?.Reason ?? string.Empty,
                        string.Join(";", entry.Aliases),
                        entry.LegacyLookupName,
                        entry.SourceKind,
                        entry.SourceDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                        entry.SourceReference,
                        price.Override?.SourceReference ?? string.Empty);
                    DataGridViewRow row = priceGrid.Rows[index];
                    row.Tag = price.Override;
                    ApplyOverrideCellState(row, price.IsOverridden);
                }
                dirty = markDirty;
            }
            finally
            {
                loading = false;
            }
            RefreshCoverageSummary(profile);
        }

        private void SwitchAudience()
        {
            if (loading)
                return;
            MachineRateAudience requested = SelectedAudience;
            if (requested == activeAudience)
                return;
            try
            {
                PriceProfile currentDraft = BuildCurrentProfile(activeAudience);
                profileDrafts[activeAudience] = currentDraft;
                PriceProfile target;
                if (!profileDrafts.TryGetValue(requested, out target))
                {
                    DialogResult answer = MessageBox.Show(
                        "Chưa có hồ sơ giá " + AudienceText(requested) + ".\n\n" +
                        "Tạo một bản nháp từ hồ sơ " + AudienceText(activeAudience) +
                        " để bạn rà soát lại giá nhân công và máy?",
                        "Tạo hồ sơ giá",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    if (answer != DialogResult.Yes)
                    {
                        SelectAudience(activeAudience);
                        return;
                    }
                    target = CloneForAudience(currentDraft, requested);
                    profileDrafts[requested] = target;
                }
                LoadProfile(target, false);
                statusLabel.ForeColor = Color.DarkGoldenrod;
                statusLabel.Text = "Đang sửa hồ sơ " + AudienceText(requested) +
                    ". Các thay đổi chỉ được dùng khi bấm Lưu hồ sơ.";
            }
            catch (Exception ex)
            {
                SelectAudience(activeAudience);
                Execute("Chuyển đối tượng lương", () => { throw ex; });
            }
        }

        private static PriceProfile CloneForAudience(
            PriceProfile source,
            MachineRateAudience audience)
        {
            string suffix = audience == MachineRateAudience.StateBudgetSalary ? "-HLNS" : "-KHLNS";
            return PriceProfile.Create(
                source.ProfileId + suffix,
                source.DataVersion,
                source.DisplayName + " - " + AudienceText(audience),
                source.Location,
                source.ValuationDate,
                audience,
                DateTime.UtcNow,
                source.Entries,
                source.Overrides);
        }

        private void ReloadLegacyProfile()
        {
            LoadProfile(LegacyWorkbookPriceProfileService.Import(workbook), true);
            statusLabel.Text = "Đã đọc lại VL-NC-M; thay đổi chưa được lưu.";
        }

        private void ImportProfile()
        {
            Directory.CreateDirectory(AppPaths.PriceProfileDirectory);
            using (var dialog = new OpenFileDialog
            {
                Filter = "TTBMVN PriceProfile (*.ttbprice)|*.ttbprice|All files (*.*)|*.*",
                InitialDirectory = AppPaths.PriceProfileDirectory,
                RestoreDirectory = true
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                PriceProfile imported = PriceProfileStore.LoadFile(dialog.FileName);
                profileDrafts[imported.LaborAudience] = imported;
                LoadProfile(imported, false);
                statusLabel.Text = "Đã import profile; bấm Lưu hồ sơ để pin vào workbook.";
            }
        }

        private void ExportProfile()
        {
            PriceProfile profile = BuildCurrentProfile();
            Directory.CreateDirectory(AppPaths.PriceProfileDirectory);
            using (var dialog = new SaveFileDialog
            {
                Filter = "TTBMVN PriceProfile (*.ttbprice)|*.ttbprice",
                InitialDirectory = AppPaths.PriceProfileDirectory,
                FileName = profile.ProfileId + "-" + profile.DataVersion + ".ttbprice",
                RestoreDirectory = true
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;
                File.WriteAllText(
                    dialog.FileName,
                    PriceProfileSerializer.Serialize(profile),
                    new UTF8Encoding(false));
                statusLabel.Text = "Đã export: " + dialog.FileName;
            }
        }

        private void ShowCoverage()
        {
            PriceProfile profile = BuildCurrentProfile();
            IReadOnlyList<PriceRequirement> requirements =
                WorkbookPriceRequirementService.LoadPinnedNormRequirements(workbook);
            PriceCoverageResult coverage = PriceProfileCoverageValidator.Validate(profile, requirements);
            if (coverage.IsComplete)
            {
                MessageBox.Show(
                    "Bảng giá đáp ứng toàn bộ mã nguồn lực của package đang pin.",
                    "Kiểm tra bảng giá",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
            string detail = string.Join(Environment.NewLine, coverage.Issues
                .Take(30)
                .Select(item => item.Message));
            if (coverage.Issues.Count > 30)
                detail += Environment.NewLine + "... và " + (coverage.Issues.Count - 30) + " lỗi khác.";
            MessageBox.Show(
                detail,
                "Bảng giá còn thiếu hoặc không tương thích",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private void RefreshCoverageSummary(PriceProfile profile)
        {
            try
            {
                IReadOnlyList<PriceRequirement> requirements =
                    WorkbookPriceRequirementService.LoadPinnedNormRequirements(workbook);
                PriceCoverageResult coverage = PriceProfileCoverageValidator.Validate(profile, requirements);
                statusLabel.ForeColor = coverage.IsComplete ? Color.DarkGreen : Color.DarkGoldenrod;
                statusLabel.Text = profile.Entries.Count + " giá; " +
                    (coverage.IsComplete
                        ? "đủ mã nguồn lực của package đang pin."
                        : "còn " + coverage.Issues.Count + " yêu cầu thiếu/sai.");
            }
            catch (Exception ex)
            {
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
            }
        }

        private void AddRow()
        {
            int index = priceGrid.Rows.Add(
                string.Empty,
                PriceResourceKind.Material,
                string.Empty,
                "each",
                0m,
                false,
                0m,
                string.Empty,
                string.Empty,
                string.Empty,
                PriceSourceKind.Manual,
                valuationDateInput.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                string.Empty,
                string.Empty);
            ApplyOverrideCellState(priceGrid.Rows[index], false);
            priceGrid.CurrentCell = priceGrid.Rows[index].Cells["colPriceCode"];
            priceGrid.BeginEdit(true);
            dirty = true;
        }

        private void DeleteSelectedRows()
        {
            foreach (DataGridViewRow row in priceGrid.SelectedRows.Cast<DataGridViewRow>().ToArray())
            {
                if (!row.IsNewRow)
                    priceGrid.Rows.Remove(row);
            }
            dirty = true;
        }

        private DataGridView CreateGrid()
        {
            var grid = new DataGridView
            {
                Name = "gridPriceProfile",
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            };
            grid.Columns.Add(TextColumn("colPriceCode", "Mã", 125));
            grid.Columns.Add(ComboColumn<PriceResourceKind>("colPriceKind", "Loại", 88, new[]
            {
                new OptionItem<PriceResourceKind>("Vật liệu", PriceResourceKind.Material),
                new OptionItem<PriceResourceKind>("Nhân công", PriceResourceKind.Labor),
                new OptionItem<PriceResourceKind>("Nhiên liệu", PriceResourceKind.FuelEnergy),
                new OptionItem<PriceResourceKind>("Giá gốc máy", PriceResourceKind.MachineOriginalPrice),
                new OptionItem<PriceResourceKind>("Ca máy", PriceResourceKind.MachineShift)
            }));
            grid.Columns.Add(TextColumn("colPriceName", "Tên dữ liệu", 180));
            grid.Columns.Add(TextColumn("colPriceUnit", "ĐVT", 62));
            grid.Columns.Add(TextColumn("colBasePrice", "Giá gốc", 92));
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "colHasOverride", HeaderText = "Ghi đè", Width = 55 });
            grid.Columns.Add(TextColumn("colAppliedPrice", "Giá áp dụng", 100));
            grid.Columns.Add(TextColumn("colOverrideReason", "Lý do ghi đè", 150));
            grid.Columns.Add(TextColumn("colAliases", "Alias (;)", 120));
            grid.Columns.Add(TextColumn("colLegacyName", "Tên lookup cũ", 165));
            grid.Columns.Add(ComboColumn<PriceSourceKind>("colSourceKind", "Loại nguồn", 105, new[]
            {
                new OptionItem<PriceSourceKind>("Nhập tay", PriceSourceKind.Manual),
                new OptionItem<PriceSourceKind>("Từ workbook", PriceSourceKind.WorkbookImport),
                new OptionItem<PriceSourceKind>("Báo giá", PriceSourceKind.MarketQuote),
                new OptionItem<PriceSourceKind>("Công bố giá", PriceSourceKind.PublishedNotice),
                new OptionItem<PriceSourceKind>("Gói pháp lý", PriceSourceKind.RegulationReference),
                new OptionItem<PriceSourceKind>("Tính toán", PriceSourceKind.Calculated)
            }));
            grid.Columns.Add(TextColumn("colSourceDate", "Ngày nguồn", 88));
            grid.Columns.Add(TextColumn("colPriceSource", "Nguồn giá", 170));
            grid.Columns.Add(TextColumn("colOverrideSource", "Nguồn ghi đè", 150));
            grid.CurrentCellDirtyStateChanged += (sender, args) =>
            {
                if (grid.IsCurrentCellDirty)
                    grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            grid.CellValueChanged += GridCellValueChanged;
            grid.RowsAdded += (sender, args) => { if (!loading) dirty = true; };
            grid.RowsRemoved += (sender, args) => { if (!loading) dirty = true; };
            grid.DataError += (sender, args) => { args.ThrowException = true; };
            return grid;
        }

        private void GridCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (loading || e.RowIndex < 0)
                return;
            dirty = true;
            DataGridViewRow row = priceGrid.Rows[e.RowIndex];
            string column = priceGrid.Columns[e.ColumnIndex].Name;
            if (column == "colHasOverride")
                ApplyOverrideCellState(row, Boolean(row, "colHasOverride"));
            else if (column == "colBasePrice" && !Boolean(row, "colHasOverride"))
                row.Cells["colAppliedPrice"].Value = row.Cells["colBasePrice"].Value;
        }

        private static void ApplyOverrideCellState(DataGridViewRow row, bool enabled)
        {
            foreach (string column in new[] { "colAppliedPrice", "colOverrideReason", "colOverrideSource" })
            {
                row.Cells[column].ReadOnly = !enabled;
                row.Cells[column].Style.BackColor = enabled ? SystemColors.Window : SystemColors.Control;
            }
            if (!enabled)
            {
                row.Cells["colAppliedPrice"].Value = row.Cells["colBasePrice"].Value;
                row.Cells["colOverrideReason"].Value = string.Empty;
                row.Cells["colOverrideSource"].Value = string.Empty;
                row.Tag = null;
            }
        }

        private void HookDirtyEvents()
        {
            profileIdInput.TextChanged += (sender, args) => MarkDirty();
            versionInput.TextChanged += (sender, args) => MarkDirty();
            displayNameInput.TextChanged += (sender, args) => MarkDirty();
            locationInput.TextChanged += (sender, args) => MarkDirty();
            valuationDateInput.ValueChanged += (sender, args) => MarkDirty();
        }

        private MachineRateAudience SelectedAudience =>
            ((OptionItem<MachineRateAudience>)audienceInput.SelectedItem).Value;

        private static string AudienceText(MachineRateAudience audience)
        {
            return audience == MachineRateAudience.StateBudgetSalary
                ? "Hưởng lương NSNN"
                : "Không hưởng lương NSNN";
        }

        private void MarkDirty()
        {
            if (!loading)
                dirty = true;
        }

        private void Execute(string operation, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                RuntimeLogger.LogOperation(
                    ex, operation, "DT-404", operation, workbook, "ResourcePrices");
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
                MessageBox.Show(ex.Message, operation, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SelectAudience(MachineRateAudience audience)
        {
            for (int index = 0; index < audienceInput.Items.Count; index++)
            {
                if (((OptionItem<MachineRateAudience>)audienceInput.Items[index]).Value == audience)
                {
                    audienceInput.SelectedIndex = index;
                    return;
                }
            }
            audienceInput.SelectedIndex = 0;
        }

        private decimal Money(DataGridViewRow row, string column)
        {
            object raw = row.Cells[column].Value;
            if (raw is decimal decimalValue)
                return decimalValue;
            decimal value;
            string text = Convert.ToString(raw, numberCulture)?.Trim();
            if (!decimal.TryParse(text, NumberStyles.Number, numberCulture, out value))
                throw new ArgumentException("Giá tại dòng " + (row.Index + 1) + " không hợp lệ.");
            return value;
        }

        private static DateTime Date(DataGridViewRow row, string column)
        {
            string text = CellText(row, column);
            DateTime value;
            if (!DateTime.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out value))
                throw new ArgumentException("Ngày nguồn tại dòng " + (row.Index + 1) + " phải theo yyyy-MM-dd.");
            return value.Date;
        }

        private static T Value<T>(DataGridViewRow row, string column) where T : struct
        {
            object raw = row.Cells[column].Value;
            if (raw is T value)
                return value;
            T parsed;
            if (Enum.TryParse(Convert.ToString(raw, CultureInfo.InvariantCulture), out parsed))
                return parsed;
            throw new ArgumentException("Giá trị tại dòng " + (row.Index + 1) + " không hợp lệ.");
        }

        private static string CellText(DataGridViewRow row, string column)
        {
            return (Convert.ToString(row.Cells[column].Value, CultureInfo.CurrentCulture) ?? string.Empty).Trim();
        }

        private static bool Boolean(DataGridViewRow row, string column)
        {
            return Convert.ToBoolean(row.Cells[column].Value ?? false, CultureInfo.InvariantCulture);
        }

        private static TextBox CreateTextBox(string name)
        {
            return new TextBox { Name = name, Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle };
        }

        private static void AddMetadata(
            TableLayoutPanel panel,
            int column,
            int row,
            string title,
            Control control)
        {
            panel.Controls.Add(new Label
            {
                Text = title,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            }, column, row);
            panel.Controls.Add(control, column, row + 1);
        }

        private static Button CreateButton(string name, string text, int width)
        {
            return new Button { Name = name, Text = text, Width = width, Height = 30 };
        }

        private static DataGridViewTextBoxColumn TextColumn(string name, string title, int width)
        {
            return new DataGridViewTextBoxColumn { Name = name, HeaderText = title, Width = width };
        }

        private static DataGridViewComboBoxColumn ComboColumn<T>(
            string name,
            string title,
            int width,
            IEnumerable<OptionItem<T>> items)
        {
            return new DataGridViewComboBoxColumn
            {
                Name = name,
                HeaderText = title,
                Width = width,
                DataSource = items.ToList(),
                DisplayMember = "Text",
                ValueMember = "Value",
                ValueType = typeof(T),
                DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
            };
        }

        private sealed class OptionItem<T>
        {
            public OptionItem(string text, T value)
            {
                Text = text;
                Value = value;
            }

            public string Text { get; }
            public T Value { get; }
            public override string ToString() => Text;
        }
    }
}
