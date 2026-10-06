namespace ExcelAddIn1.Winform
{
    partial class FrmDaodat
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.pnlBottom = new System.Windows.Forms.Panel();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnSupport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.pnlFill = new System.Windows.Forms.Panel();
            this.grpSheets = new System.Windows.Forms.GroupBox();
            this.cboSheets = new System.Windows.Forms.ComboBox();
            this.txtNameSheet = new System.Windows.Forms.TextBox();
            this.lblNameSheet = new System.Windows.Forms.Label();
            this.radSheetHienCo = new System.Windows.Forms.RadioButton();
            this.radNewSheet = new System.Windows.Forms.RadioButton();
            this.chkLinkBack = new System.Windows.Forms.CheckBox();
            this.pnlTableData = new System.Windows.Forms.Panel();
            this.txtCotKetQuaDaoDap = new System.Windows.Forms.TextBox();
            this.lblCotKetQuaDaoDap = new System.Windows.Forms.Label();
            this.chkUseDefaultOutputRow = new System.Windows.Forms.CheckBox();
            this.txtHangTrongDaoDap = new System.Windows.Forms.TextBox();
            this.lblHangTrongDaoDap = new System.Windows.Forms.Label();
            this.dgvTableData = new System.Windows.Forms.DataGridView();
            this.colTenDuLieu = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCotTrongData = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCotTrongDaoDap = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDinhDang = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnChonTableData = new System.Windows.Forms.Button();
            this.txtTableData = new System.Windows.Forms.TextBox();
            this.lblTableData = new System.Windows.Forms.Label();
            this.chkInsertRow = new System.Windows.Forms.CheckBox();
            this.grpLoaiDuAn = new System.Windows.Forms.GroupBox();
            this.radDuAnCa3m5m = new System.Windows.Forms.RadioButton();
            this.radDuAnChi3m = new System.Windows.Forms.RadioButton();
            this.txtMaxAttempts = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.lblMaxAttempts = new System.Windows.Forms.Label();
            this.cboMaxWorkers = new System.Windows.Forms.ComboBox();
            this.lblMaxWorkers = new System.Windows.Forms.Label();
            this.lblVolumeRange = new System.Windows.Forms.Label();
            this.tableParams = new System.Windows.Forms.TableLayoutPanel();
            this.txtVMax = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.txtVMin = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.lblTheTichV = new System.Windows.Forms.Label();
            this.lblParam = new System.Windows.Forms.Label();
            this.lblMin = new System.Windows.Forms.Label();
            this.lblMax = new System.Windows.Forms.Label();
            this.lblD1 = new System.Windows.Forms.Label();
            this.txtD1Min = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.txtD1Max = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.lblR1 = new System.Windows.Forms.Label();
            this.txtR1Min = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.txtR1Max = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.lblD2 = new System.Windows.Forms.Label();
            this.txtD2Min = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.txtD2Max = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.lblR2 = new System.Windows.Forms.Label();
            this.txtR2Min = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.txtR2Max = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.lblH = new System.Windows.Forms.Label();
            this.txtHMin = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.txtHMax = new ExcelAddIn1.Winform.SmartNumericUpDown();
            this.grpType = new System.Windows.Forms.GroupBox();
            this.rad5m = new System.Windows.Forms.RadioButton();
            this.rad3m = new System.Windows.Forms.RadioButton();
            this.pnlBottom.SuspendLayout();
            this.pnlFill.SuspendLayout();
            this.grpSheets.SuspendLayout();
            this.pnlTableData.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTableData)).BeginInit();
            this.grpLoaiDuAn.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtMaxAttempts)).BeginInit();
            this.tableParams.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtVMax)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtVMin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtD1Min)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtD1Max)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtR1Min)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtR1Max)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtD2Min)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtD2Max)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtR2Min)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtR2Max)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtHMin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtHMax)).BeginInit();
            this.grpType.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlBottom
            // 
            this.pnlBottom.Controls.Add(this.btnClose);
            this.pnlBottom.Controls.Add(this.btnRun);
            this.pnlBottom.Controls.Add(this.btnSave);
            this.pnlBottom.Controls.Add(this.btnSupport);
            this.pnlBottom.Controls.Add(this.btnImport);
            this.pnlBottom.Controls.Add(this.btnExport);
            this.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlBottom.Location = new System.Drawing.Point(0, 554);
            this.pnlBottom.Name = "pnlBottom";
            this.pnlBottom.Size = new System.Drawing.Size(624, 66);
            this.pnlBottom.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(522, 18);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(90, 32);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "Đóng";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnRun
            // 
            this.btnRun.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRun.Location = new System.Drawing.Point(410, 18);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(106, 32);
            this.btnRun.TabIndex = 1;
            this.btnRun.Text = "Lưu và chạy";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Location = new System.Drawing.Point(298, 18);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(106, 32);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "Lưu cài đặt";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnSupport
            // 
            this.btnSupport.Location = new System.Drawing.Point(212, 18);
            this.btnSupport.Name = "btnSupport";
            this.btnSupport.Size = new System.Drawing.Size(80, 32);
            this.btnSupport.TabIndex = 5;
            this.btnSupport.Text = "Hỗ trợ";
            this.btnSupport.UseVisualStyleBackColor = true;
            this.btnSupport.Click += new System.EventHandler(this.btnSupport_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(112, 18);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(90, 32);
            this.btnImport.TabIndex = 4;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(12, 18);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(90, 32);
            this.btnExport.TabIndex = 3;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // pnlFill
            // 
            this.pnlFill.Controls.Add(this.grpSheets);
            this.pnlFill.Controls.Add(this.pnlTableData);
            this.pnlFill.Controls.Add(this.grpLoaiDuAn);
            this.pnlFill.Controls.Add(this.txtMaxAttempts);
            this.pnlFill.Controls.Add(this.lblMaxAttempts);
            this.pnlFill.Controls.Add(this.cboMaxWorkers);
            this.pnlFill.Controls.Add(this.lblMaxWorkers);
            this.pnlFill.Controls.Add(this.lblVolumeRange);
            this.pnlFill.Controls.Add(this.tableParams);
            this.pnlFill.Controls.Add(this.grpType);
            this.pnlFill.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFill.Location = new System.Drawing.Point(0, 0);
            this.pnlFill.Name = "pnlFill";
            this.pnlFill.Padding = new System.Windows.Forms.Padding(12);
            this.pnlFill.Size = new System.Drawing.Size(624, 554);
            this.pnlFill.TabIndex = 1;
            // 
            // grpSheets
            // 
            this.grpSheets.Controls.Add(this.cboSheets);
            this.grpSheets.Controls.Add(this.txtNameSheet);
            this.grpSheets.Controls.Add(this.lblNameSheet);
            this.grpSheets.Controls.Add(this.radSheetHienCo);
            this.grpSheets.Controls.Add(this.radNewSheet);
            this.grpSheets.Controls.Add(this.chkLinkBack);
            this.grpSheets.Location = new System.Drawing.Point(15, 322);
            this.grpSheets.Name = "grpSheets";
            this.grpSheets.Size = new System.Drawing.Size(258, 120);
            this.grpSheets.TabIndex = 3;
            this.grpSheets.TabStop = false;
            this.grpSheets.Text = "Sheets";
            // 
            // cboSheets
            // 
            this.cboSheets.FormattingEnabled = true;
            this.cboSheets.Location = new System.Drawing.Point(114, 79);
            this.cboSheets.Name = "cboSheets";
            this.cboSheets.Size = new System.Drawing.Size(121, 24);
            this.cboSheets.TabIndex = 16;
            // 
            // txtNameSheet
            // 
            this.txtNameSheet.Location = new System.Drawing.Point(114, 50);
            this.txtNameSheet.Name = "txtNameSheet";
            this.txtNameSheet.Size = new System.Drawing.Size(121, 22);
            this.txtNameSheet.TabIndex = 15;
            // 
            // lblNameSheet
            // 
            this.lblNameSheet.AutoSize = true;
            this.lblNameSheet.Location = new System.Drawing.Point(15, 53);
            this.lblNameSheet.Name = "lblNameSheet";
            this.lblNameSheet.Size = new System.Drawing.Size(80, 16);
            this.lblNameSheet.TabIndex = 14;
            this.lblNameSheet.Text = "Name sheet";
            // 
            // radSheetHienCo
            // 
            this.radSheetHienCo.AutoSize = true;
            this.radSheetHienCo.Location = new System.Drawing.Point(132, 24);
            this.radSheetHienCo.Name = "radSheetHienCo";
            this.radSheetHienCo.Size = new System.Drawing.Size(109, 20);
            this.radSheetHienCo.TabIndex = 1;
            this.radSheetHienCo.Text = "Sheet hiện có";
            this.radSheetHienCo.UseVisualStyleBackColor = true;
            this.radSheetHienCo.CheckedChanged += new System.EventHandler(this.radSheetMode_CheckedChanged);
            // 
            // radNewSheet
            // 
            this.radNewSheet.AutoSize = true;
            this.radNewSheet.Checked = true;
            this.radNewSheet.Location = new System.Drawing.Point(18, 24);
            this.radNewSheet.Name = "radNewSheet";
            this.radNewSheet.Size = new System.Drawing.Size(91, 20);
            this.radNewSheet.TabIndex = 0;
            this.radNewSheet.TabStop = true;
            this.radNewSheet.Text = "New sheet";
            this.radNewSheet.UseVisualStyleBackColor = true;
            this.radNewSheet.CheckedChanged += new System.EventHandler(this.radSheetMode_CheckedChanged);
            // 
            // chkLinkBack
            // 
            this.chkLinkBack.AutoSize = true;
            this.chkLinkBack.Enabled = false;
            this.chkLinkBack.Location = new System.Drawing.Point(15, 79);
            this.chkLinkBack.Name = "chkLinkBack";
            this.chkLinkBack.Size = new System.Drawing.Size(93, 20);
            this.chkLinkBack.TabIndex = 5;
            this.chkLinkBack.Text = "Link ngược";
            this.chkLinkBack.UseVisualStyleBackColor = true;
            // 
            // pnlTableData
            // 
            this.pnlTableData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlTableData.Controls.Add(this.txtCotKetQuaDaoDap);
            this.pnlTableData.Controls.Add(this.lblCotKetQuaDaoDap);
            this.pnlTableData.Controls.Add(this.chkUseDefaultOutputRow);
            this.pnlTableData.Controls.Add(this.txtHangTrongDaoDap);
            this.pnlTableData.Controls.Add(this.lblHangTrongDaoDap);
            this.pnlTableData.Controls.Add(this.dgvTableData);
            this.pnlTableData.Controls.Add(this.btnChonTableData);
            this.pnlTableData.Controls.Add(this.txtTableData);
            this.pnlTableData.Controls.Add(this.lblTableData);
            this.pnlTableData.Controls.Add(this.chkInsertRow);
            this.pnlTableData.Location = new System.Drawing.Point(280, 79);
            this.pnlTableData.Name = "pnlTableData";
            this.pnlTableData.Size = new System.Drawing.Size(332, 422);
            this.pnlTableData.TabIndex = 13;
            // 
            // txtCotKetQuaDaoDap
            // 
            this.txtCotKetQuaDaoDap.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCotKetQuaDaoDap.Location = new System.Drawing.Point(155, 86);
            this.txtCotKetQuaDaoDap.Name = "txtCotKetQuaDaoDap";
            this.txtCotKetQuaDaoDap.Size = new System.Drawing.Size(162, 22);
            this.txtCotKetQuaDaoDap.TabIndex = 11;
            // 
            // lblCotKetQuaDaoDap
            // 
            this.lblCotKetQuaDaoDap.AutoSize = true;
            this.lblCotKetQuaDaoDap.Location = new System.Drawing.Point(5, 89);
            this.lblCotKetQuaDaoDap.Name = "lblCotKetQuaDaoDap";
            this.lblCotKetQuaDaoDap.Size = new System.Drawing.Size(92, 16);
            this.lblCotKetQuaDaoDap.TabIndex = 10;
            this.lblCotKetQuaDaoDap.Text = "TH3;V3;TH5;V5";
            // 
            // txtHangTrongDaoDap
            // 
            this.txtHangTrongDaoDap.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtHangTrongDaoDap.Location = new System.Drawing.Point(155, 59);
            this.txtHangTrongDaoDap.Name = "txtHangTrongDaoDap";
            this.txtHangTrongDaoDap.Size = new System.Drawing.Size(162, 22);
            this.txtHangTrongDaoDap.TabIndex = 9;
            // 
            // chkUseDefaultOutputRow
            // 
            this.chkUseDefaultOutputRow.AutoSize = true;
            this.chkUseDefaultOutputRow.Checked = true;
            this.chkUseDefaultOutputRow.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUseDefaultOutputRow.Location = new System.Drawing.Point(5, 59);
            this.chkUseDefaultOutputRow.Name = "chkUseDefaultOutputRow";
            this.chkUseDefaultOutputRow.Size = new System.Drawing.Size(128, 20);
            this.chkUseDefaultOutputRow.TabIndex = 12;
            this.chkUseDefaultOutputRow.Text = "Hàng mặc định";
            this.chkUseDefaultOutputRow.UseVisualStyleBackColor = true;
            this.chkUseDefaultOutputRow.CheckedChanged += new System.EventHandler(this.chkUseDefaultOutputRow_CheckedChanged);
            // 
            // lblHangTrongDaoDap
            // 
            this.lblHangTrongDaoDap.AutoSize = true;
            this.lblHangTrongDaoDap.Location = new System.Drawing.Point(5, 61);
            this.lblHangTrongDaoDap.Name = "lblHangTrongDaoDap";
            this.lblHangTrongDaoDap.Size = new System.Drawing.Size(127, 16);
            this.lblHangTrongDaoDap.TabIndex = 8;
            this.lblHangTrongDaoDap.Text = "Hàng trong đào đắp";
            this.lblHangTrongDaoDap.Visible = false;
            // 
            // dgvTableData
            // 
            this.dgvTableData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvTableData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTableData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colTenDuLieu,
            this.colCotTrongData,
            this.colCotTrongDaoDap,
            this.colDinhDang});
            this.dgvTableData.Location = new System.Drawing.Point(-1, 135);
            this.dgvTableData.Name = "dgvTableData";
            this.dgvTableData.RowHeadersVisible = false;
            this.dgvTableData.RowHeadersWidth = 51;
            this.dgvTableData.RowTemplate.Height = 24;
            this.dgvTableData.Size = new System.Drawing.Size(330, 281);
            this.dgvTableData.TabIndex = 7;
            // 
            // colTenDuLieu
            // 
            this.colTenDuLieu.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colTenDuLieu.HeaderText = "Tên dữ liệu";
            this.colTenDuLieu.MinimumWidth = 6;
            this.colTenDuLieu.Name = "colTenDuLieu";
            // 
            // colCotTrongData
            // 
            this.colCotTrongData.HeaderText = "Cột trong Data";
            this.colCotTrongData.MinimumWidth = 6;
            this.colCotTrongData.Name = "colCotTrongData";
            this.colCotTrongData.Width = 40;
            // 
            // colCotTrongDaoDap
            // 
            this.colCotTrongDaoDap.HeaderText = "Cột trong đào đắp";
            this.colCotTrongDaoDap.MinimumWidth = 6;
            this.colCotTrongDaoDap.Name = "colCotTrongDaoDap";
            this.colCotTrongDaoDap.Width = 40;
            // 
            // colDinhDang
            // 
            this.colDinhDang.HeaderText = "Định dạng";
            this.colDinhDang.MinimumWidth = 6;
            this.colDinhDang.Name = "colDinhDang";
            this.colDinhDang.Width = 70;
            // 
            // btnChonTableData
            // 
            this.btnChonTableData.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnChonTableData.Location = new System.Drawing.Point(256, 30);
            this.btnChonTableData.Name = "btnChonTableData";
            this.btnChonTableData.Size = new System.Drawing.Size(63, 23);
            this.btnChonTableData.TabIndex = 6;
            this.btnChonTableData.Text = "Chọn";
            this.btnChonTableData.UseVisualStyleBackColor = true;
            // 
            // txtTableData
            // 
            this.txtTableData.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTableData.Location = new System.Drawing.Point(87, 30);
            this.txtTableData.Name = "txtTableData";
            this.txtTableData.Size = new System.Drawing.Size(162, 22);
            this.txtTableData.TabIndex = 5;
            // 
            // lblTableData
            // 
            this.lblTableData.AutoSize = true;
            this.lblTableData.Location = new System.Drawing.Point(5, 33);
            this.lblTableData.Name = "lblTableData";
            this.lblTableData.Size = new System.Drawing.Size(75, 16);
            this.lblTableData.TabIndex = 4;
            this.lblTableData.Text = "Table Data";
            // 
            // chkInsertRow
            // 
            this.chkInsertRow.AutoSize = true;
            this.chkInsertRow.Location = new System.Drawing.Point(5, 6);
            this.chkInsertRow.Name = "chkInsertRow";
            this.chkInsertRow.Size = new System.Drawing.Size(101, 20);
            this.chkInsertRow.TabIndex = 3;
            this.chkInsertRow.Text = "Xuống dòng";
            this.chkInsertRow.UseVisualStyleBackColor = true;
            // 
            // grpLoaiDuAn
            // 
            this.grpLoaiDuAn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpLoaiDuAn.Controls.Add(this.radDuAnCa3m5m);
            this.grpLoaiDuAn.Controls.Add(this.radDuAnChi3m);
            this.grpLoaiDuAn.Location = new System.Drawing.Point(279, 15);
            this.grpLoaiDuAn.Name = "grpLoaiDuAn";
            this.grpLoaiDuAn.Size = new System.Drawing.Size(333, 58);
            this.grpLoaiDuAn.TabIndex = 2;
            this.grpLoaiDuAn.TabStop = false;
            this.grpLoaiDuAn.Text = "Loại dự án";
            // 
            // radDuAnCa3m5m
            // 
            this.radDuAnCa3m5m.AutoSize = true;
            this.radDuAnCa3m5m.Location = new System.Drawing.Point(132, 24);
            this.radDuAnCa3m5m.Name = "radDuAnCa3m5m";
            this.radDuAnCa3m5m.Size = new System.Drawing.Size(105, 20);
            this.radDuAnCa3m5m.TabIndex = 1;
            this.radDuAnCa3m5m.Text = "Cả 3m và 5m";
            this.radDuAnCa3m5m.UseVisualStyleBackColor = true;
            // 
            // radDuAnChi3m
            // 
            this.radDuAnChi3m.AutoSize = true;
            this.radDuAnChi3m.Checked = true;
            this.radDuAnChi3m.Location = new System.Drawing.Point(18, 24);
            this.radDuAnChi3m.Name = "radDuAnChi3m";
            this.radDuAnChi3m.Size = new System.Drawing.Size(86, 20);
            this.radDuAnChi3m.TabIndex = 0;
            this.radDuAnChi3m.TabStop = true;
            this.radDuAnChi3m.Text = "Chỉ có 3m";
            this.radDuAnChi3m.UseVisualStyleBackColor = true;
            // 
            // txtMaxAttempts
            // 
            this.txtMaxAttempts.Location = new System.Drawing.Point(129, 448);
            this.txtMaxAttempts.Name = "txtMaxAttempts";
            this.txtMaxAttempts.Size = new System.Drawing.Size(78, 22);
            this.txtMaxAttempts.TabIndex = 7;
            // 
            // lblMaxAttempts
            // 
            this.lblMaxAttempts.AutoSize = true;
            this.lblMaxAttempts.Location = new System.Drawing.Point(12, 451);
            this.lblMaxAttempts.Name = "lblMaxAttempts";
            this.lblMaxAttempts.Size = new System.Drawing.Size(101, 16);
            this.lblMaxAttempts.TabIndex = 6;
            this.lblMaxAttempts.Text = "Số lần thử tối đa";
            // 
            // cboMaxWorkers
            // 
            this.cboMaxWorkers.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMaxWorkers.FormattingEnabled = true;
            this.cboMaxWorkers.Location = new System.Drawing.Point(129, 476);
            this.cboMaxWorkers.Name = "cboMaxWorkers";
            this.cboMaxWorkers.Size = new System.Drawing.Size(110, 24);
            this.cboMaxWorkers.TabIndex = 11;
            // 
            // lblMaxWorkers
            // 
            this.lblMaxWorkers.AutoSize = true;
            this.lblMaxWorkers.Location = new System.Drawing.Point(12, 479);
            this.lblMaxWorkers.Name = "lblMaxWorkers";
            this.lblMaxWorkers.Size = new System.Drawing.Size(108, 16);
            this.lblMaxWorkers.TabIndex = 12;
            this.lblMaxWorkers.Text = "Số luồng tối đa";
            // 
            // lblVolumeRange
            // 
            this.lblVolumeRange.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblVolumeRange.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(248)))), ((int)(((byte)(252)))));
            this.lblVolumeRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblVolumeRange.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblVolumeRange.Location = new System.Drawing.Point(15, 504);
            this.lblVolumeRange.Name = "lblVolumeRange";
            this.lblVolumeRange.Size = new System.Drawing.Size(594, 40);
            this.lblVolumeRange.TabIndex = 2;
            this.lblVolumeRange.Text = "V min - max:";
            this.lblVolumeRange.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // tableParams
            // 
            this.tableParams.ColumnCount = 3;
            this.tableParams.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 102F));
            this.tableParams.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableParams.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableParams.Controls.Add(this.txtVMax, 2, 6);
            this.tableParams.Controls.Add(this.txtVMin, 1, 6);
            this.tableParams.Controls.Add(this.lblTheTichV, 0, 6);
            this.tableParams.Controls.Add(this.lblParam, 0, 0);
            this.tableParams.Controls.Add(this.lblMin, 1, 0);
            this.tableParams.Controls.Add(this.lblMax, 2, 0);
            this.tableParams.Controls.Add(this.lblD1, 0, 1);
            this.tableParams.Controls.Add(this.txtD1Min, 1, 1);
            this.tableParams.Controls.Add(this.txtD1Max, 2, 1);
            this.tableParams.Controls.Add(this.lblR1, 0, 2);
            this.tableParams.Controls.Add(this.txtR1Min, 1, 2);
            this.tableParams.Controls.Add(this.txtR1Max, 2, 2);
            this.tableParams.Controls.Add(this.lblD2, 0, 3);
            this.tableParams.Controls.Add(this.txtD2Min, 1, 3);
            this.tableParams.Controls.Add(this.txtD2Max, 2, 3);
            this.tableParams.Controls.Add(this.lblR2, 0, 4);
            this.tableParams.Controls.Add(this.txtR2Min, 1, 4);
            this.tableParams.Controls.Add(this.txtR2Max, 2, 4);
            this.tableParams.Controls.Add(this.lblH, 0, 5);
            this.tableParams.Controls.Add(this.txtHMin, 1, 5);
            this.tableParams.Controls.Add(this.txtHMax, 2, 5);
            this.tableParams.Location = new System.Drawing.Point(15, 79);
            this.tableParams.Name = "tableParams";
            this.tableParams.RowCount = 6;
            this.tableParams.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableParams.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableParams.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableParams.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableParams.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableParams.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.tableParams.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableParams.Size = new System.Drawing.Size(258, 231);
            this.tableParams.TabIndex = 1;
            // 
            // txtVMax
            // 
            this.txtVMax.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtVMax.Location = new System.Drawing.Point(183, 203);
            this.txtVMax.Name = "txtVMax";
            this.txtVMax.Size = new System.Drawing.Size(72, 22);
            this.txtVMax.TabIndex = 20;
            // 
            // txtVMin
            // 
            this.txtVMin.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtVMin.Location = new System.Drawing.Point(105, 203);
            this.txtVMin.Name = "txtVMin";
            this.txtVMin.Size = new System.Drawing.Size(72, 22);
            this.txtVMin.TabIndex = 19;
            // 
            // lblTheTichV
            // 
            this.lblTheTichV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTheTichV.Location = new System.Drawing.Point(3, 200);
            this.lblTheTichV.Name = "lblTheTichV";
            this.lblTheTichV.Size = new System.Drawing.Size(96, 31);
            this.lblTheTichV.TabIndex = 18;
            this.lblTheTichV.Text = "Thể tích (V)";
            this.lblTheTichV.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblParam
            // 
            this.lblParam.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblParam.Location = new System.Drawing.Point(3, 0);
            this.lblParam.Name = "lblParam";
            this.lblParam.Size = new System.Drawing.Size(96, 30);
            this.lblParam.TabIndex = 0;
            this.lblParam.Text = "Thông số";
            this.lblParam.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblMin
            // 
            this.lblMin.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblMin.Location = new System.Drawing.Point(105, 0);
            this.lblMin.Name = "lblMin";
            this.lblMin.Size = new System.Drawing.Size(72, 30);
            this.lblMin.TabIndex = 1;
            this.lblMin.Text = "Min";
            this.lblMin.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblMax
            // 
            this.lblMax.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblMax.Location = new System.Drawing.Point(183, 0);
            this.lblMax.Name = "lblMax";
            this.lblMax.Size = new System.Drawing.Size(72, 30);
            this.lblMax.TabIndex = 2;
            this.lblMax.Text = "Max";
            this.lblMax.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblD1
            // 
            this.lblD1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblD1.Location = new System.Drawing.Point(3, 30);
            this.lblD1.Name = "lblD1";
            this.lblD1.Size = new System.Drawing.Size(96, 34);
            this.lblD1.TabIndex = 3;
            this.lblD1.Text = "Dài (d1)";
            this.lblD1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtD1Min
            // 
            this.txtD1Min.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtD1Min.Location = new System.Drawing.Point(105, 33);
            this.txtD1Min.Name = "txtD1Min";
            this.txtD1Min.Size = new System.Drawing.Size(72, 22);
            this.txtD1Min.TabIndex = 4;
            // 
            // txtD1Max
            // 
            this.txtD1Max.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtD1Max.Location = new System.Drawing.Point(183, 33);
            this.txtD1Max.Name = "txtD1Max";
            this.txtD1Max.Size = new System.Drawing.Size(72, 22);
            this.txtD1Max.TabIndex = 5;
            // 
            // lblR1
            // 
            this.lblR1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblR1.Location = new System.Drawing.Point(3, 64);
            this.lblR1.Name = "lblR1";
            this.lblR1.Size = new System.Drawing.Size(96, 34);
            this.lblR1.TabIndex = 6;
            this.lblR1.Text = "Rộng (r1)";
            this.lblR1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtR1Min
            // 
            this.txtR1Min.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtR1Min.Location = new System.Drawing.Point(105, 67);
            this.txtR1Min.Name = "txtR1Min";
            this.txtR1Min.Size = new System.Drawing.Size(72, 22);
            this.txtR1Min.TabIndex = 7;
            // 
            // txtR1Max
            // 
            this.txtR1Max.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtR1Max.Location = new System.Drawing.Point(183, 67);
            this.txtR1Max.Name = "txtR1Max";
            this.txtR1Max.Size = new System.Drawing.Size(72, 22);
            this.txtR1Max.TabIndex = 8;
            // 
            // lblD2
            // 
            this.lblD2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblD2.Location = new System.Drawing.Point(3, 98);
            this.lblD2.Name = "lblD2";
            this.lblD2.Size = new System.Drawing.Size(96, 34);
            this.lblD2.TabIndex = 9;
            this.lblD2.Text = "Dài (d2)";
            this.lblD2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtD2Min
            // 
            this.txtD2Min.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtD2Min.Location = new System.Drawing.Point(105, 101);
            this.txtD2Min.Name = "txtD2Min";
            this.txtD2Min.Size = new System.Drawing.Size(72, 22);
            this.txtD2Min.TabIndex = 10;
            // 
            // txtD2Max
            // 
            this.txtD2Max.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtD2Max.Location = new System.Drawing.Point(183, 101);
            this.txtD2Max.Name = "txtD2Max";
            this.txtD2Max.Size = new System.Drawing.Size(72, 22);
            this.txtD2Max.TabIndex = 11;
            // 
            // lblR2
            // 
            this.lblR2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblR2.Location = new System.Drawing.Point(3, 132);
            this.lblR2.Name = "lblR2";
            this.lblR2.Size = new System.Drawing.Size(96, 34);
            this.lblR2.TabIndex = 12;
            this.lblR2.Text = "Rộng (r2)";
            this.lblR2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtR2Min
            // 
            this.txtR2Min.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtR2Min.Location = new System.Drawing.Point(105, 135);
            this.txtR2Min.Name = "txtR2Min";
            this.txtR2Min.Size = new System.Drawing.Size(72, 22);
            this.txtR2Min.TabIndex = 13;
            // 
            // txtR2Max
            // 
            this.txtR2Max.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtR2Max.Location = new System.Drawing.Point(183, 135);
            this.txtR2Max.Name = "txtR2Max";
            this.txtR2Max.Size = new System.Drawing.Size(72, 22);
            this.txtR2Max.TabIndex = 14;
            // 
            // lblH
            // 
            this.lblH.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblH.Location = new System.Drawing.Point(3, 166);
            this.lblH.Name = "lblH";
            this.lblH.Size = new System.Drawing.Size(96, 34);
            this.lblH.TabIndex = 15;
            this.lblH.Text = "Độ sâu hố (H)";
            this.lblH.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtHMin
            // 
            this.txtHMin.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtHMin.Location = new System.Drawing.Point(105, 169);
            this.txtHMin.Name = "txtHMin";
            this.txtHMin.Size = new System.Drawing.Size(72, 22);
            this.txtHMin.TabIndex = 16;
            // 
            // txtHMax
            // 
            this.txtHMax.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtHMax.Location = new System.Drawing.Point(183, 169);
            this.txtHMax.Name = "txtHMax";
            this.txtHMax.Size = new System.Drawing.Size(72, 22);
            this.txtHMax.TabIndex = 17;
            // 
            // grpType
            // 
            this.grpType.Controls.Add(this.rad5m);
            this.grpType.Controls.Add(this.rad3m);
            this.grpType.Location = new System.Drawing.Point(15, 15);
            this.grpType.Name = "grpType";
            this.grpType.Size = new System.Drawing.Size(258, 58);
            this.grpType.TabIndex = 0;
            this.grpType.TabStop = false;
            this.grpType.Text = "Loại hố đào";
            // 
            // rad5m
            // 
            this.rad5m.AutoSize = true;
            this.rad5m.Location = new System.Drawing.Point(132, 24);
            this.rad5m.Name = "rad5m";
            this.rad5m.Size = new System.Drawing.Size(94, 20);
            this.rad5m.TabIndex = 1;
            this.rad5m.Text = "Hố đào 5m";
            this.rad5m.UseVisualStyleBackColor = true;
            this.rad5m.CheckedChanged += new System.EventHandler(this.radType_CheckedChanged);
            // 
            // rad3m
            // 
            this.rad3m.AutoSize = true;
            this.rad3m.Checked = true;
            this.rad3m.Location = new System.Drawing.Point(18, 24);
            this.rad3m.Name = "rad3m";
            this.rad3m.Size = new System.Drawing.Size(94, 20);
            this.rad3m.TabIndex = 0;
            this.rad3m.TabStop = true;
            this.rad3m.Text = "Hố đào 3m";
            this.rad3m.UseVisualStyleBackColor = true;
            this.rad3m.CheckedChanged += new System.EventHandler(this.radType_CheckedChanged);
            // 
            // FrmDaodat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(624, 620);
            this.Controls.Add(this.pnlFill);
            this.Controls.Add(this.pnlBottom);
            this.MinimumSize = new System.Drawing.Size(640, 507);
            this.Name = "FrmDaodat";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Xử lý hố đào";
            this.pnlBottom.ResumeLayout(false);
            this.pnlFill.ResumeLayout(false);
            this.pnlFill.PerformLayout();
            this.grpSheets.ResumeLayout(false);
            this.grpSheets.PerformLayout();
            this.pnlTableData.ResumeLayout(false);
            this.pnlTableData.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTableData)).EndInit();
            this.grpLoaiDuAn.ResumeLayout(false);
            this.grpLoaiDuAn.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtMaxAttempts)).EndInit();
            this.tableParams.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txtVMax)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtVMin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtD1Min)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtD1Max)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtR1Min)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtR1Max)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtD2Min)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtD2Max)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtR2Min)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtR2Max)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtHMin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtHMax)).EndInit();
            this.grpType.ResumeLayout(false);
            this.grpType.PerformLayout();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.Panel pnlBottom;
        private System.Windows.Forms.Panel pnlFill;
        private System.Windows.Forms.GroupBox grpType;
        private System.Windows.Forms.RadioButton rad5m;
        private System.Windows.Forms.RadioButton rad3m;
        private System.Windows.Forms.TableLayoutPanel tableParams;
        private System.Windows.Forms.Label lblParam;
        private System.Windows.Forms.Label lblMin;
        private System.Windows.Forms.Label lblMax;
        private System.Windows.Forms.Label lblD1;
        private SmartNumericUpDown txtD1Min;
        private SmartNumericUpDown txtD1Max;
        private System.Windows.Forms.Label lblR1;
        private SmartNumericUpDown txtR1Min;
        private SmartNumericUpDown txtR1Max;
        private System.Windows.Forms.Label lblD2;
        private SmartNumericUpDown txtD2Min;
        private SmartNumericUpDown txtD2Max;
        private System.Windows.Forms.Label lblR2;
        private SmartNumericUpDown txtR2Min;
        private SmartNumericUpDown txtR2Max;
        private System.Windows.Forms.Label lblH;
        private SmartNumericUpDown txtHMax;
        private System.Windows.Forms.Label lblVolumeRange;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnSupport;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.CheckBox chkInsertRow;
        private System.Windows.Forms.CheckBox chkLinkBack;
        private System.Windows.Forms.Label lblMaxAttempts;
        private SmartNumericUpDown txtMaxAttempts;
        private System.Windows.Forms.Label lblMaxWorkers;
        private System.Windows.Forms.ComboBox cboMaxWorkers;
        private SmartNumericUpDown txtVMax;
        private SmartNumericUpDown txtVMin;
        private System.Windows.Forms.Label lblTheTichV;
        private SmartNumericUpDown txtHMin;
        private System.Windows.Forms.GroupBox grpLoaiDuAn;
        private System.Windows.Forms.RadioButton radDuAnCa3m5m;
        private System.Windows.Forms.RadioButton radDuAnChi3m;
        private System.Windows.Forms.Panel pnlTableData;
        private System.Windows.Forms.Button btnChonTableData;
        private System.Windows.Forms.TextBox txtTableData;
        private System.Windows.Forms.Label lblTableData;
        private System.Windows.Forms.DataGridView dgvTableData;
        private System.Windows.Forms.GroupBox grpSheets;
        private System.Windows.Forms.RadioButton radSheetHienCo;
        private System.Windows.Forms.RadioButton radNewSheet;
        private System.Windows.Forms.TextBox txtHangTrongDaoDap;
        private System.Windows.Forms.Label lblHangTrongDaoDap;
        private System.Windows.Forms.CheckBox chkUseDefaultOutputRow;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTenDuLieu;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCotTrongData;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCotTrongDaoDap;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDinhDang;
        private System.Windows.Forms.ComboBox cboSheets;
        private System.Windows.Forms.TextBox txtNameSheet;
        private System.Windows.Forms.Label lblNameSheet;
        private System.Windows.Forms.TextBox txtCotKetQuaDaoDap;
        private System.Windows.Forms.Label lblCotKetQuaDaoDap;
    }
}

