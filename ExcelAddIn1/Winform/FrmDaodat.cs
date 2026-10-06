using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public partial class FrmDaodat : Form, IWorkbookSheetListObserver
    {
        private enum VolumeHeightSource
        {
            Volume,
            Height
        }

        private bool isLoading;
        private bool isSyncingVolumeHeight;
        private bool isChoosingExcelFormat;
        private DateTime skipWorkbookSheetRefreshUntil = DateTime.MinValue;
        private readonly Excel.Workbook sourceWorkbook;
        private IDisposable sheetListSubscription;
        private VolumeHeightSource volumeHeightSource = VolumeHeightSource.Volume;
        private static readonly string[] DefaultDataKeys =
        {
            "TH3", "V3", "D1_3", "R1_3", "D2_3", "R2_3", "H_3",
            "TH5", "V5", "D1_5", "R1_5", "D2_5", "R2_5", "H_5"
        };
        private const string KeySerialNumber = "key_STT";

        public FrmDaodat()
            : this(false)
        {
        }

        public FrmDaodat(bool insertRow)
        {
            ExcelCulture.ApplyToCurrentThread(Globals.ThisAddIn.Application);
            sourceWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            InitializeComponent();
            ConfigureNumericControls();
            ConfigureTableDataGrid();
            chkInsertRow.Checked = true;
            chkUseDefaultOutputRow.Checked = true;
            txtCotKetQuaDaoDap.Text = "A;B;C;D";
            RegisterTextChangedEvents();
            LoadWorkbookSheets();
            EnsureDefaultGridRows();
            LoadSettingsToForm();
            btnChonTableData.Click += btnChonTableData_Click;
            cboSheets.DropDown += cboSheets_DropDown;
            Activated += FrmDaodat_Activated;
            FormClosed += FrmDaodat_FormClosed;
            if (sourceWorkbook != null && Globals.ThisAddIn.SheetChangeCoordinator != null)
                sheetListSubscription = Globals.ThisAddIn.SheetChangeCoordinator.Subscribe(sourceWorkbook, this);
        }

        private bool Is5m => rad5m.Checked;

        private void RegisterTextChangedEvents()
        {
            foreach (SmartNumericUpDown control in GetDimensionControls())
            {
                control.TextChanged += txtDimensionRange_TextChanged;
            }

            txtHMin.TextChanged += txtHeightRange_TextChanged;
            txtHMax.TextChanged += txtHeightRange_TextChanged;
            txtVMin.TextChanged += txtVolumeRange_TextChanged;
            txtVMax.TextChanged += txtVolumeRange_TextChanged;
            txtMaxAttempts.TextChanged += txtOtherNumber_TextChanged;
            cboMaxWorkers.SelectedIndexChanged += txtOtherNumber_TextChanged;
        }

        private IEnumerable<SmartNumericUpDown> GetDimensionControls()
        {
            yield return txtD1Min;
            yield return txtD1Max;
            yield return txtR1Min;
            yield return txtR1Max;
            yield return txtD2Min;
            yield return txtD2Max;
            yield return txtR2Min;
            yield return txtR2Max;
        }

        private void ConfigureNumericControls()
        {
            foreach (SmartNumericUpDown control in GetDimensionControls())
                ConfigureDecimalNumeric(control, 2);

            ConfigureDecimalNumeric(txtHMin, 4);
            ConfigureDecimalNumeric(txtHMax, 4);
            ConfigureDecimalNumeric(txtVMin, 2);
            ConfigureDecimalNumeric(txtVMax, 2);
            ConfigureIntegerNumeric(txtMaxAttempts);
            ConfigureMaxWorkersCombo();
        }

        private void ConfigureDecimalNumeric(SmartNumericUpDown control, int maximumDecimalPlaces)
        {
            control.Minimum = 0;
            control.Maximum = 1000000000;
            control.MaximumDecimalPlaces = maximumDecimalPlaces;
            control.DecimalPlaces = 0;
            control.Increment = 1;
        }

        private void ConfigureIntegerNumeric(SmartNumericUpDown control)
        {
            control.Minimum = 0;
            control.Maximum = int.MaxValue;
            control.MaximumDecimalPlaces = 0;
            control.DecimalPlaces = 0;
            control.Increment = 1;
        }

        private void ConfigureMaxWorkersCombo()
        {
            cboMaxWorkers.Items.Clear();
            cboMaxWorkers.Items.Add("0 - Tự động");

            int maxWorkers = GetMachineMaxWorkers();
            for (int i = 1; i <= maxWorkers; i++)
                cboMaxWorkers.Items.Add(i.ToString(CultureInfo.CurrentCulture));

            cboMaxWorkers.SelectedIndex = 0;
        }

        private void ConfigureTableDataGrid()
        {
            dgvTableData.AllowUserToDeleteRows = true;
            dgvTableData.CellBeginEdit += dgvTableData_CellBeginEdit;
            dgvTableData.CellValueChanged += dgvTableData_CellValueChanged;
            dgvTableData.RowsAdded += dgvTableData_RowsAdded;
            dgvTableData.UserDeletingRow += dgvTableData_UserDeletingRow;
            dgvTableData.KeyDown += dgvTableData_KeyDown;
            dgvTableData.CellDoubleClick += dgvTableData_CellDoubleClick;
        }

        private void LoadSettingsToForm()
        {
            isLoading = true;
            ApplyParameterSettings(RandomDaodat.GetSettings(Is5m));
            ApplyRunOptions(RandomDaodat.GetRunOptions(Is5m));
            volumeHeightSource = VolumeHeightSource.Volume;

            isLoading = false;
            UpdateHeightFromVolume();
        }

        private void LoadParameterSettingsToForm()
        {
            isLoading = true;
            ApplyParameterSettings(RandomDaodat.GetSettings(Is5m));
            volumeHeightSource = VolumeHeightSource.Volume;

            isLoading = false;
            UpdateHeightFromVolume();
        }

        private void ApplyParameterSettings(DaodatRandomSettings settings)
        {
            SetNumericValue(txtD1Min, settings.D1Min);
            SetNumericValue(txtD1Max, settings.D1Max);
            SetNumericValue(txtR1Min, settings.R1Min);
            SetNumericValue(txtR1Max, settings.R1Max);
            SetNumericValue(txtD2Min, settings.D2Min);
            SetNumericValue(txtD2Max, settings.D2Max);
            SetNumericValue(txtR2Min, settings.R2Min);
            SetNumericValue(txtR2Max, settings.R2Max);
            SetNumericValue(txtHMin, settings.HMin);
            SetNumericValue(txtHMax, settings.HMax);
            SetNumericValue(txtVMin, settings.VolumeMin);
            SetNumericValue(txtVMax, settings.VolumeMax);
            SetIntegerValue(txtMaxAttempts, settings.MaxAttempts);
            SetMaxWorkersValue(settings.MaxParallelWorkers);
        }

        private void ApplyRunOptions(DaodatRunOptions runOptions)
        {
            chkInsertRow.Checked = runOptions.InsertRows;
            radDuAnCa3m5m.Checked = runOptions.ProjectIncludes5m;
            radDuAnChi3m.Checked = !runOptions.ProjectIncludes5m;
            radNewSheet.Checked = runOptions.CreateNewSheet;
            radSheetHienCo.Checked = !runOptions.CreateNewSheet;
            chkLinkBack.Enabled = true;
            chkLinkBack.Checked = runOptions.LinkBack;
            txtNameSheet.Text = string.IsNullOrWhiteSpace(runOptions.NewSheetName) ? "HoDao" : runOptions.NewSheetName;
            if (!string.IsNullOrWhiteSpace(runOptions.ExistingSheetName))
                cboSheets.Text = runOptions.ExistingSheetName;
            txtTableData.Text = runOptions.TableDataAddress ?? string.Empty;
            chkUseDefaultOutputRow.Checked = runOptions.UseDefaultOutputRow;
            txtHangTrongDaoDap.Text = runOptions.OutputStartRow > 0
                ? runOptions.OutputStartRow.ToString(CultureInfo.CurrentCulture)
                : txtHangTrongDaoDap.Text;
            txtCotKetQuaDaoDap.Text = FormatSourceDataColumns(runOptions.SourceDataColumns);
            ApplyGridMappings(runOptions.ColumnMappings);
            ToggleSheetControls();
            ToggleOutputRowControls();
        }

        private DaodatRandomSettings ReadSettingsFromForm()
        {
            var settings = new DaodatRandomSettings
            {
                D1Min = ParseNumber(txtD1Min, "can duoi d1"),
                D1Max = ParseNumber(txtD1Max, "can tren d1"),
                R1Min = ParseNumber(txtR1Min, "can duoi r1"),
                R1Max = ParseNumber(txtR1Max, "can tren r1"),
                D2Min = ParseNumber(txtD2Min, "can duoi d2"),
                D2Max = ParseNumber(txtD2Max, "can tren d2"),
                R2Min = ParseNumber(txtR2Min, "can duoi r2"),
                R2Max = ParseNumber(txtR2Max, "can tren r2"),
                HMin = ParseNumber(txtHMin, "can duoi H"),
                HMax = ParseNumber(txtHMax, "can tren H"),
                MaxAttempts = ParsePositiveInt(txtMaxAttempts, "so lan thu toi da"),
                MaxParallelWorkers = ParseMaxWorkers()
            };

            settings.Validate();
            return settings;
        }

        private void SaveSettings()
        {
            RandomDaodat.SaveSettings(Is5m, ReadSettingsFromForm());
            RandomDaodat.SaveRunOptions(BuildRunOptions());
            RefreshCalculatedFieldsFromActiveSource();
        }

        private DaodatRunOptions BuildRunOptions()
        {
            bool hasTableData = !string.IsNullOrWhiteSpace(txtTableData.Text);
            var runOptions = new DaodatRunOptions
            {
                Is5m = Is5m,
                InsertRows = true,
                CreateNewSheet = radNewSheet.Checked,
                LinkBack = chkLinkBack.Checked,
                UseOutputStartCell = false,
                OutputStartAddress = null,
                HasTableData = hasTableData,
                ProjectIncludes5m = radDuAnCa3m5m.Checked,
                UseDefaultOutputRow = chkUseDefaultOutputRow.Checked,
                TableDataAddress = txtTableData.Text.Trim(),
                NewSheetName = txtNameSheet.Text.Trim(),
                ExistingSheetName = cboSheets.Text.Trim(),
                SourceDataColumns = ParseSourceDataColumns(),
                ColumnMappings = ReadGridMappings()
            };

            if (hasTableData && runOptions.UseDefaultOutputRow)
                runOptions.OutputStartRow = ParsePositiveInt(txtHangTrongDaoDap, "hang trong dao dap");

            return runOptions;
        }

        private void RefreshCalculatedFieldsFromActiveSource()
        {
            if (volumeHeightSource == VolumeHeightSource.Height)
                UpdateVolumeFromHeight();
            else
                UpdateHeightFromVolume();
        }

        private void UpdateVolumeFromHeight()
        {
            if (isLoading || isSyncingVolumeHeight)
                return;

            try
            {
                DaodatRandomSettings settings = ReadSettingsFromForm();
                isSyncingVolumeHeight = true;
                SetNumericValue(txtVMin, settings.VolumeMin);
                SetNumericValue(txtVMax, settings.VolumeMax);
                isSyncingVolumeHeight = false;

                UpdateVolumeLabel(settings.VolumeMin, settings.VolumeMax);
            }
            catch (Exception ex)
            {
                isSyncingVolumeHeight = false;
                ShowVolumeError(ex);
            }
        }

        private void UpdateHeightFromVolume()
        {
            if (isLoading || isSyncingVolumeHeight)
                return;

            try
            {
                double d1Min = ParseNumber(txtD1Min, "can duoi d1");
                double r1Min = ParseNumber(txtR1Min, "can duoi r1");
                double d2Min = ParseNumber(txtD2Min, "can duoi d2");
                double r2Min = ParseNumber(txtR2Min, "can duoi r2");
                double d1Max = ParseNumber(txtD1Max, "can tren d1");
                double r1Max = ParseNumber(txtR1Max, "can tren r1");
                double d2Max = ParseNumber(txtD2Max, "can tren d2");
                double r2Max = ParseNumber(txtR2Max, "can tren r2");
                double vMin = ParseNumber(txtVMin, "can duoi V");
                double vMax = ParseNumber(txtVMax, "can tren V");

                isSyncingVolumeHeight = true;
                SetNumericValue(txtHMin, CalculateHeightFromVolume(vMin, d1Min, r1Min, d2Min, r2Min));
                SetNumericValue(txtHMax, CalculateHeightFromVolume(vMax, d1Max, r1Max, d2Max, r2Max));
                isSyncingVolumeHeight = false;

                DaodatRandomSettings settings = ReadSettingsFromForm();
                UpdateVolumeLabel(settings.VolumeMin, settings.VolumeMax);
            }
            catch (Exception ex)
            {
                isSyncingVolumeHeight = false;
                ShowVolumeError(ex);
            }
        }

        private void UpdateVolumeLabelOnly()
        {
            if (isLoading || isSyncingVolumeHeight)
                return;

            try
            {
                DaodatRandomSettings settings = ReadSettingsFromForm();
                UpdateVolumeLabel(settings.VolumeMin, settings.VolumeMax);
            }
            catch (Exception ex)
            {
                ShowVolumeError(ex);
            }
        }

        private void UpdateVolumeLabel(double volumeMin, double volumeMax)
        {
            lblVolumeRange.ForeColor = Color.FromArgb(20, 84, 45);
            lblVolumeRange.Text = $"V co the tao: {volumeMin:0.###} den {volumeMax:0.###}";
        }

        private void ShowVolumeError(Exception ex)
        {
            lblVolumeRange.ForeColor = Color.FromArgb(180, 30, 30);
            lblVolumeRange.Text = ex.Message;
        }

        private double CalculateHeightFromVolume(double volume, double d1, double r1, double d2, double r2)
        {
            return DaodatMath.CalculateHeightFromVolume(volume, d1, r1, d2, r2);
        }

        private void SetNumericValue(SmartNumericUpDown control, double value)
        {
            decimal roundedValue = RoundForControl(control, Convert.ToDecimal(value));
            string text = FormatNumber((double)roundedValue, control.MaximumDecimalPlaces);
            control.SetSmartValue(roundedValue, text);
        }

        private void SetIntegerValue(SmartNumericUpDown control, int value)
        {
            control.SetSmartValue(value, value.ToString(CultureInfo.CurrentCulture));
        }

        private void SetMaxWorkersValue(int value)
        {
            value = Math.Max(0, Math.Min(value, GetMachineMaxWorkers()));
            string prefix = value.ToString(CultureInfo.CurrentCulture);

            for (int i = 0; i < cboMaxWorkers.Items.Count; i++)
            {
                string itemText = Convert.ToString(cboMaxWorkers.Items[i]);
                if (itemText == prefix || itemText.StartsWith(prefix + " ", StringComparison.Ordinal))
                {
                    cboMaxWorkers.SelectedIndex = i;
                    return;
                }
            }

            cboMaxWorkers.SelectedIndex = 0;
        }

        private double ParseNumber(SmartNumericUpDown control, string fieldName)
        {
            string text = (control.Text ?? string.Empty).Trim();

            if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double value))
            {
                return value;
            }

            throw new ArgumentException($"Gia tri {fieldName} khong hop le.");
        }

        private int ParsePositiveInt(Control control, string fieldName)
        {
            string text = (control.Text ?? string.Empty).Trim();
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out int value) && value > 0)
                return value;

            throw new ArgumentException($"Gia tri {fieldName} phai la so nguyen lon hon 0.");
        }

        private int ParseNonNegativeInt(Control control, string fieldName)
        {
            string text = (control.Text ?? string.Empty).Trim();
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out int value) && value >= 0)
                return value;

            throw new ArgumentException($"Gia tri {fieldName} phai la so nguyen >= 0. Nhap 0 de tu dong.");
        }

        private int ParseMaxWorkers()
        {
            string text = (cboMaxWorkers.Text ?? string.Empty).Trim();
            int separatorIndex = text.IndexOf(' ');
            if (separatorIndex > 0)
                text = text.Substring(0, separatorIndex);

            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out int value))
                return 0;

            return Math.Max(0, Math.Min(value, GetMachineMaxWorkers()));
        }

        private int GetMachineMaxWorkers()
        {
            return Math.Max(1, Environment.ProcessorCount);
        }

        private string FormatNumber(double value)
        {
            return FormatNumber(value, 4);
        }

        private string FormatNumber(double value, int maximumDecimalPlaces)
        {
            return value.ToString("0." + new string('#', maximumDecimalPlaces), CultureInfo.CurrentCulture);
        }

        private decimal RoundForControl(SmartNumericUpDown control, decimal value)
        {
            int decimals = Math.Max(0, control.MaximumDecimalPlaces);
            return Math.Round(value, decimals, MidpointRounding.AwayFromZero);
        }

        private void radType_CheckedChanged(object sender, EventArgs e)
        {
            if (!isLoading && ((RadioButton)sender).Checked)
                LoadParameterSettingsToForm();
        }

        private void txtDimensionRange_TextChanged(object sender, EventArgs e)
        {
            RefreshCalculatedFieldsFromActiveSource();
        }

        private void txtHeightRange_TextChanged(object sender, EventArgs e)
        {
            if (isLoading || isSyncingVolumeHeight)
                return;

            volumeHeightSource = VolumeHeightSource.Height;
            UpdateVolumeFromHeight();
        }

        private void txtVolumeRange_TextChanged(object sender, EventArgs e)
        {
            if (isLoading || isSyncingVolumeHeight)
                return;

            volumeHeightSource = VolumeHeightSource.Volume;
            UpdateHeightFromVolume();
        }

        private void txtOtherNumber_TextChanged(object sender, EventArgs e)
        {
            UpdateVolumeLabelOnly();
        }

        private void radSheetMode_CheckedChanged(object sender, EventArgs e)
        {
            chkLinkBack.Enabled = true;
            ToggleSheetControls();
        }

        private void chkUseDefaultOutputRow_CheckedChanged(object sender, EventArgs e)
        {
            ToggleOutputRowControls();
        }

        private void ToggleSheetControls()
        {
            txtNameSheet.Enabled = radNewSheet.Checked;
            cboSheets.Enabled = radSheetHienCo.Checked;
        }

        private void ToggleOutputRowControls()
        {
            txtHangTrongDaoDap.Enabled = chkUseDefaultOutputRow.Checked;
        }

        private void LoadWorkbookSheets()
        {
            LoadWorkbookSheets(preserveCurrentSelection: false);
        }

        private void LoadWorkbookSheets(bool preserveCurrentSelection)
        {
            if (sourceWorkbook == null)
            {
                ApplyWorkbookSheets(new WorkbookSheetDescriptor[0], preserveCurrentSelection);
                return;
            }

            ApplyWorkbookSheets(
                WorkbookSheetChangeCoordinator.CaptureSnapshot(sourceWorkbook),
                preserveCurrentSelection);
        }

        public void OnWorkbookSheetsChanged(IReadOnlyList<WorkbookSheetDescriptor> sheets)
        {
            if (IsDisposed || Disposing || isChoosingExcelFormat ||
                DateTime.Now < skipWorkbookSheetRefreshUntil)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action<IReadOnlyList<WorkbookSheetDescriptor>>(OnWorkbookSheetsChanged), sheets);
                return;
            }

            ApplyWorkbookSheets(sheets, preserveCurrentSelection: true);
        }

        private void ApplyWorkbookSheets(
            IEnumerable<WorkbookSheetDescriptor> sheets,
            bool preserveCurrentSelection)
        {
            string previousKey = preserveCurrentSelection
                ? (cboSheets.SelectedItem as WorkbookSheetDescriptor)?.Key
                : string.Empty;
            string previousName = preserveCurrentSelection ? cboSheets.Text : string.Empty;
            IReadOnlyList<WorkbookSheetDescriptor> sheetList = WorkbookSheetList.Copy(sheets);
            WorkbookSheetDescriptor selected = WorkbookSheetList.ResolveSelection(
                sheetList,
                previousKey,
                previousName);

            cboSheets.BeginUpdate();
            try
            {
                cboSheets.Items.Clear();
                foreach (WorkbookSheetDescriptor sheet in sheetList)
                    cboSheets.Items.Add(sheet);

                if (selected != null)
                    cboSheets.SelectedItem = selected;
                else
                    cboSheets.Text = string.Empty;
            }
            finally
            {
                cboSheets.EndUpdate();
            }
        }

        private void RefreshWorkbookSheets()
        {
            if (isChoosingExcelFormat || DateTime.Now < skipWorkbookSheetRefreshUntil)
                return;

            LoadWorkbookSheets(preserveCurrentSelection: true);
        }

        private void cboSheets_DropDown(object sender, EventArgs e)
        {
            RefreshWorkbookSheets();
        }

        private void FrmDaodat_Activated(object sender, EventArgs e)
        {
            RefreshWorkbookSheets();
        }

        private void FrmDaodat_FormClosed(object sender, FormClosedEventArgs e)
        {
            sheetListSubscription?.Dispose();
            sheetListSubscription = null;
        }

        private void EnsureDefaultGridRows()
        {
            if (dgvTableData.Rows.Count > 1)
            {
                ApplyGridRowRules();
                return;
            }

            AddDefaultGridRows();
            ApplyGridRowRules();
        }

        private void AddDefaultGridRows()
        {
            foreach (string key in DefaultDataKeys)
                dgvTableData.Rows.Add(key, string.Empty, string.Empty, GetDefaultFormatLocal(key, string.Empty));
        }

        private void ApplyGridMappings(List<DaodatColumnMapping> mappings)
        {
            dgvTableData.Rows.Clear();

            var mappingByKey = new Dictionary<string, DaodatColumnMapping>(StringComparer.OrdinalIgnoreCase);
            var customMappings = new List<DaodatColumnMapping>();
            foreach (DaodatColumnMapping mapping in mappings ?? new List<DaodatColumnMapping>())
            {
                string key = NormalizeGridKey(mapping.Key);
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                if (IsDefaultDataKey(key))
                    mappingByKey[key] = mapping;
                else
                    customMappings.Add(mapping);
            }

            foreach (string key in DefaultDataKeys)
            {
                if (mappingByKey.TryGetValue(key, out DaodatColumnMapping mapping))
                {
                    dgvTableData.Rows.Add(
                        key,
                        mapping.SourceColumn ?? string.Empty,
                        mapping.OutputColumn ?? string.Empty,
                        ResolveFormatLocal(key, mapping.SourceColumn, mapping.FormatLocal));
                }
                else
                {
                    dgvTableData.Rows.Add(key, string.Empty, string.Empty, GetDefaultFormatLocal(key, string.Empty));
                }
            }

            foreach (DaodatColumnMapping mapping in customMappings)
            {
                dgvTableData.Rows.Add(
                    mapping.Key ?? string.Empty,
                    mapping.SourceColumn ?? string.Empty,
                    mapping.OutputColumn ?? string.Empty,
                    ResolveFormatLocal(mapping.Key, mapping.SourceColumn, mapping.FormatLocal));
            }

            ApplyGridRowRules();
        }

        private void ApplyGridRowRules()
        {
            foreach (DataGridViewRow row in dgvTableData.Rows)
            {
                if (row.IsNewRow)
                    continue;

                bool isDefaultRow = IsDefaultGridRow(row);
                row.Cells[colTenDuLieu.Name].ReadOnly = isDefaultRow;
            }
        }

        private bool IsDefaultGridRow(DataGridViewRow row)
        {
            return row != null && IsDefaultDataKey(Convert.ToString(row.Cells[colTenDuLieu.Name].Value));
        }

        private bool IsDefaultDataKey(string key)
        {
            key = NormalizeGridKey(key);
            foreach (string defaultKey in DefaultDataKeys)
            {
                if (string.Equals(defaultKey, key, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private string NormalizeGridKey(string key)
        {
            return (key ?? string.Empty).Trim().ToUpperInvariant();
        }

        private void dgvTableData_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            DataGridViewRow row = dgvTableData.Rows[e.RowIndex];
            if (IsDefaultGridRow(row) && dgvTableData.Columns[e.ColumnIndex].Name == colTenDuLieu.Name)
                e.Cancel = true;
        }

        private void dgvTableData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                ApplyGridRowRules();
        }

        private void dgvTableData_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            ApplyGridRowRules();
        }

        private void dgvTableData_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            if (IsDefaultGridRow(e.Row))
                e.Cancel = true;
        }

        private void dgvTableData_KeyDown(object sender, KeyEventArgs e)
        {
            if (dgvTableData.IsCurrentCellInEditMode)
                return;

            if (e.KeyCode == Keys.Delete)
            {
                HandleGridDeleteKey();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Back)
            {
                HandleGridBackspaceKey();
                e.Handled = true;
            }
        }

        private void HandleGridDeleteKey()
        {
            if (dgvTableData.SelectedRows.Count > 0)
            {
                var rows = new List<DataGridViewRow>();
                foreach (DataGridViewRow row in dgvTableData.SelectedRows)
                    rows.Add(row);

                foreach (DataGridViewRow row in rows)
                {
                    if (row.IsNewRow)
                        continue;

                    if (IsDefaultGridRow(row))
                    {
                        row.Cells[colCotTrongData.Name].Value = null;
                        row.Cells[colCotTrongDaoDap.Name].Value = null;
                        row.Cells[colDinhDang.Name].Value = GetDefaultFormatLocal(
                            Convert.ToString(row.Cells[colTenDuLieu.Name].Value),
                            string.Empty);
                    }
                    else
                    {
                        dgvTableData.Rows.Remove(row);
                    }
                }

                return;
            }

            foreach (DataGridViewCell cell in dgvTableData.SelectedCells)
            {
                if (cell.RowIndex < 0 || cell.ColumnIndex < 0)
                    continue;

                DataGridViewRow row = dgvTableData.Rows[cell.RowIndex];
                if (row.IsNewRow)
                    continue;

                bool isDefaultNameCell = IsDefaultGridRow(row) && dgvTableData.Columns[cell.ColumnIndex].Name == colTenDuLieu.Name;
                if (!isDefaultNameCell)
                    cell.Value = null;
            }
        }

        private void dgvTableData_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            if (dgvTableData.Columns[e.ColumnIndex].Name != colDinhDang.Name)
                return;

            DataGridViewRow row = dgvTableData.Rows[e.RowIndex];
            if (row.IsNewRow)
                return;

            string key = Convert.ToString(row.Cells[colTenDuLieu.Name].Value);
            string sourceColumn = Convert.ToString(row.Cells[colCotTrongData.Name].Value);
            string currentFormat = Convert.ToString(row.Cells[colDinhDang.Name].Value);
            string fallbackFormat = ResolveFormatLocal(key, sourceColumn, currentFormat);
            string selectedFormat = ShowExcelNumberFormatDialog(fallbackFormat);
            if (!string.IsNullOrWhiteSpace(selectedFormat))
                row.Cells[colDinhDang.Name].Value = selectedFormat;

            RestoreSingleGridCellSelection(e.RowIndex, e.ColumnIndex);
        }

        private void RestoreSingleGridCellSelection(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || columnIndex < 0 || rowIndex >= dgvTableData.Rows.Count || columnIndex >= dgvTableData.Columns.Count)
                return;

            DataGridViewRow row = dgvTableData.Rows[rowIndex];
            if (row.IsNewRow)
                return;

            dgvTableData.ClearSelection();
            dgvTableData.CurrentCell = row.Cells[columnIndex];
            row.Cells[columnIndex].Selected = true;
        }

        private void HandleGridBackspaceKey()
        {
            var rows = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in dgvTableData.SelectedRows)
                rows.Add(row);

            if (rows.Count == 0 && dgvTableData.CurrentRow != null)
                rows.Add(dgvTableData.CurrentRow);

            foreach (DataGridViewRow row in rows)
            {
                if (row.IsNewRow || IsDefaultGridRow(row))
                    continue;

                dgvTableData.Rows.Remove(row);
            }
        }

        private string GetSettingsDialogDirectory()
        {
            string directory = Properties.Settings.Default.DaodatSettingsDirectory;
            if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                return directory;

            directory = Path.GetDirectoryName(typeof(FrmDaodat).Assembly.Location);
            if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                return directory;

            return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        private void SaveSettingsDialogDirectory(string filePath)
        {
            string directory = Path.GetDirectoryName(filePath);
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
                return;

            Properties.Settings.Default.DaodatSettingsDirectory = directory;
            Properties.Settings.Default.Save();
        }

        private string FormatSourceDataColumns(Dictionary<string, string> columns)
        {
            if (columns == null || columns.Count == 0)
                return txtCotKetQuaDaoDap.Text;

            var parts = new List<string>();
            AddSourceDataColumnPart(parts, columns, "TH3");
            AddSourceDataColumnPart(parts, columns, "V3");
            AddSourceDataColumnPart(parts, columns, "TH5");
            AddSourceDataColumnPart(parts, columns, "V5");
            return parts.Count == 0 ? txtCotKetQuaDaoDap.Text : string.Join(";", parts);
        }

        private void AddSourceDataColumnPart(List<string> parts, Dictionary<string, string> columns, string key)
        {
            if (columns.TryGetValue(key, out string column) && !string.IsNullOrWhiteSpace(column))
                parts.Add(column);
        }

        private Dictionary<string, string> ParseSourceDataColumns()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string text = txtCotKetQuaDaoDap.Text ?? string.Empty;
            string[] items = text.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            bool hasNamedFormat = text.Contains("=");

            if (!hasNamedFormat)
            {
                string[] keys = { "TH3", "V3", "TH5", "V5" };
                if (items.Length < 2)
                    throw new ArgumentException("Dinh dang cot data phai la A;B hoac A;B;C;D.");

                if (radDuAnCa3m5m.Checked && items.Length < 4)
                    throw new ArgumentException("Du an ca 3m va 5m can khai bao cot data theo dang A;B;C;D.");

                int max = Math.Min(items.Length, keys.Length);
                for (int i = 0; i < max; i++)
                {
                    string column = NormalizeOptionalColumn(items[i], "cot data " + keys[i]);
                    if (string.IsNullOrWhiteSpace(column))
                        throw new ArgumentException("Dinh dang cot data khong hop le.");

                    result[keys[i]] = column;
                }

                return result;
            }

            foreach (string item in items)
            {
                string[] parts = item.Split('=');
                if (parts.Length != 2)
                    throw new ArgumentException("Dinh dang cot data phai la A;B;C;D.");

                string key = NormalizeDataColumnKey(parts[0]);
                string column = NormalizeOptionalColumn(parts[1], "cot data " + parts[0].Trim());
                if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(column))
                    throw new ArgumentException("Dinh dang cot data khong hop le.");

                result[key] = column;
            }

            if (!result.ContainsKey("TH3") || !result.ContainsKey("V3"))
                throw new ArgumentException("Can khai bao toi thieu Th3m va V3m trong cot data.");

            if (radDuAnCa3m5m.Checked && (!result.ContainsKey("TH5") || !result.ContainsKey("V5")))
                throw new ArgumentException("Du an ca 3m va 5m can khai bao Th5m va V5m trong cot data.");

            return result;
        }

        private string NormalizeDataColumnKey(string value)
        {
            string key = (value ?? string.Empty).Trim().ToUpperInvariant();
            if (key == "TH3M")
                return "TH3";
            if (key == "V3M")
                return "V3";
            if (key == "TH5M")
                return "TH5";
            if (key == "V5M")
                return "V5";
            return key;
        }

        private List<DaodatColumnMapping> ReadGridMappings()
        {
            var mappings = new List<DaodatColumnMapping>();
            foreach (DataGridViewRow row in dgvTableData.Rows)
            {
                if (row.IsNewRow)
                    continue;

                string key = Convert.ToString(row.Cells[colTenDuLieu.Name].Value)?.Trim();
                string sourceColumn = NormalizeGridSourceColumn(
                    Convert.ToString(row.Cells[colCotTrongData.Name].Value),
                    "cot trong Data cua " + key);
                string outputColumn = NormalizeOptionalColumn(
                    Convert.ToString(row.Cells[colCotTrongDaoDap.Name].Value),
                    "cot trong dao dap cua " + key);
                string formatLocal = Convert.ToString(row.Cells[colDinhDang.Name].Value)?.Trim();

                if (string.IsNullOrWhiteSpace(key))
                    continue;

                mappings.Add(new DaodatColumnMapping
                {
                    Key = key,
                    SourceColumn = sourceColumn,
                    OutputColumn = outputColumn,
                    FormatLocal = ResolveFormatLocal(key, sourceColumn, formatLocal)
                });
            }

            return mappings;
        }

        private string NormalizeGridSourceColumn(string value, string fieldName)
        {
            string text = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            if (IsSerialNumberSourceKey(text))
                return KeySerialNumber;

            return NormalizeOptionalColumn(text, fieldName);
        }

        private bool IsSerialNumberSourceKey(string value)
        {
            return string.Equals((value ?? string.Empty).Trim(), KeySerialNumber, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeOptionalColumn(string value, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            try
            {
                return ExcelColumnAddress.Normalize(value);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"{fieldName} khong hop le. {ex.Message}");
            }
        }

        private string ResolveFormatLocal(string key, string sourceColumn, string formatLocal)
        {
            return string.IsNullOrWhiteSpace(formatLocal)
                ? GetDefaultFormatLocal(key, sourceColumn)
                : formatLocal.Trim();
        }

        private string GetDefaultFormatLocal(string key, string sourceColumn)
        {
            if (IsSerialNumberSourceKey(sourceColumn))
                return "0";

            string normalizedKey = NormalizeGridKey(key);
            switch (normalizedKey)
            {
                case "TH3":
                case "TH5":
                    return "0";
                case "V3":
                case "V5":
                case "D1_3":
                case "R1_3":
                case "D2_3":
                case "R2_3":
                case "D1_5":
                case "R1_5":
                case "D2_5":
                case "R2_5":
                    return DecimalFormatLocal(2);
                case "H_3":
                case "H_5":
                    return DecimalFormatLocal(4);
                default:
                    return "General";
            }
        }

        private string DecimalFormatLocal(int decimalPlaces)
        {
            if (decimalPlaces <= 0)
                return "0";

            return "0" + CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + new string('0', decimalPlaces);
        }

        private string ShowExcelNumberFormatDialog(string currentFormat)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
                throw new InvalidOperationException("Chua co workbook Excel dang mo.");

            Excel.Worksheet previousSheet = excelApp.ActiveSheet as Excel.Worksheet;
            Excel.Range previousSelection = excelApp.Selection as Excel.Range;
            Excel.Worksheet tempSheet = null;
            bool oldDisplayAlerts = excelApp.DisplayAlerts;
            isChoosingExcelFormat = true;

            try
            {
                Excel.Worksheet insertAfter = previousSheet ?? workbook.Worksheets[workbook.Worksheets.Count] as Excel.Worksheet;
                tempSheet = workbook.Worksheets.Add(After: insertAfter) as Excel.Worksheet;
                tempSheet.Name = "__FormatTemp_" + DateTime.Now.ToString("HHmmss");

                Excel.Range tempCell = tempSheet.Range["A1"] as Excel.Range;
                tempCell.Value2 = 1;
                try
                {
                    tempCell.NumberFormatLocal = string.IsNullOrWhiteSpace(currentFormat) ? "General" : currentFormat;
                }
                catch
                {
                    tempCell.NumberFormatLocal = "General";
                }

                tempSheet.Activate();
                tempCell.Select();

                bool accepted = false;
                object result = excelApp.Dialogs[Excel.XlBuiltInDialog.xlDialogFormatNumber].Show();
                if (result is bool boolResult)
                    accepted = boolResult;

                return accepted ? Convert.ToString(tempCell.NumberFormatLocal) : null;
            }
            finally
            {
                try
                {
                    if (tempSheet != null)
                    {
                        excelApp.DisplayAlerts = false;
                        try
                        {
                            tempSheet.Delete();
                        }
                        catch
                        {
                            // The temporary sheet is best-effort cleanup; do not leave the form in format mode.
                        }
                    }
                }
                finally
                {
                    excelApp.DisplayAlerts = oldDisplayAlerts;
                }

                try
                {
                    previousSheet?.Activate();
                    previousSelection?.Select();
                }
                catch
                {
                    // Restoring the user's previous selection is best-effort only.
                }

                isChoosingExcelFormat = false;
                skipWorkbookSheetRefreshUntil = DateTime.Now.AddSeconds(1);
            }
        }

        private void btnChonTableData_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            object input = excelApp.InputBox(
                "Chon range Table Data:",
                "Table Data",
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

            if (input is Excel.Range range)
            {
                Excel.Worksheet sheet = range.Worksheet as Excel.Worksheet;
                txtTableData.Text = $"'{sheet.Name}'!{range.get_Address(false, false, Excel.XlReferenceStyle.xlA1)}";
            }
        }

        private bool TrySelectOutputStartCell(DaodatRunOptions runOptions)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            object input = excelApp.InputBox(
                "Chon cell bat dau do ket qua. App se lay hang cua cell nay.",
                "Chon hang trong dao dap",
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

            if (!(input is Excel.Range range))
                return false;

            Excel.Range firstCell = range.Cells[1, 1] as Excel.Range;
            if (firstCell == null)
                return false;

            Excel.Worksheet sheet = firstCell.Worksheet as Excel.Worksheet;
            runOptions.OutputStartRow = firstCell.Row;
            txtHangTrongDaoDap.Text = firstCell.Row.ToString(CultureInfo.CurrentCulture);

            if (!runOptions.CreateNewSheet && sheet != null)
            {
                runOptions.ExistingSheetName = sheet.Name;
                cboSheets.Text = sheet.Name;
            }

            return true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                SaveSettings();
                MessageBox.Show("Da luu cai dat ho dao.", "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Cai dat khong hop le", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                SaveSettings();
                using (var dialog = new SaveFileDialog())
                {
                    dialog.Title = "Xuat cai dat ho dao";
                    dialog.Filter = "Ho dao settings (*.daodatsettings)|*.daodatsettings|All files (*.*)|*.*";
                    dialog.FileName = "HoDao.daodatsettings";
                    dialog.InitialDirectory = GetSettingsDialogDirectory();
                    dialog.OverwritePrompt = true;

                    if (dialog.ShowDialog(this) != DialogResult.OK)
                        return;

                    RandomDaodat.ExportSettingsToFile(dialog.FileName);
                    SaveSettingsDialogDirectory(dialog.FileName);
                    MessageBox.Show("Da xuat cai dat ho dao.", "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Khong xuat duoc cai dat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                using (var dialog = new OpenFileDialog())
                {
                    dialog.Title = "Nhap cai dat ho dao";
                    dialog.Filter = "Ho dao settings (*.daodatsettings)|*.daodatsettings|All files (*.*)|*.*";
                    dialog.InitialDirectory = GetSettingsDialogDirectory();
                    dialog.CheckFileExists = true;

                    if (dialog.ShowDialog(this) != DialogResult.OK)
                        return;

                    RandomDaodat.ImportSettingsFromFile(dialog.FileName);
                    SaveSettingsDialogDirectory(dialog.FileName);
                    LoadSettingsToForm();
                    MessageBox.Show("Da nhap cai dat ho dao.", "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Khong nhap duoc cai dat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnSupport_Click(object sender, EventArgs e)
        {
            using (var form = new FrmSupport())
                form.ShowDialog(this);
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                DaodatRandomSettings settings = ReadSettingsFromForm();
                DaodatRunOptions runOptions = BuildRunOptions();
                if (runOptions.HasTableData && !runOptions.UseDefaultOutputRow)
                {
                    if (!TrySelectOutputStartCell(runOptions))
                        return;
                }

                RandomDaodat.SaveSettings(Is5m, settings);
                RandomDaodat.SaveRunOptions(runOptions);
                RefreshCalculatedFieldsFromActiveSource();

                DaodatRunPreview preview = RandomDaodat.BuildRunPreview(runOptions);
                if (!preview.CanRun)
                {
                    MessageBox.Show(preview.Message, "Preview khong hop le", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult confirm = MessageBox.Show(
                    preview.Message + Environment.NewLine + Environment.NewLine + "Tiep tuc chay va ghi Excel?",
                    "Preview ho dao",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button2);

                if (confirm != DialogResult.Yes)
                    return;

                RandomDaodat.Daodat(runOptions);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Cai dat khong hop le", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

