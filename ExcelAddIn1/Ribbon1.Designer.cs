namespace ExcelAddIn1
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.box2 = this.Factory.CreateRibbonBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.box3 = this.Factory.CreateRibbonBox();
            this.EdCottinhtoan = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.chkRownew = this.Factory.CreateRibbonCheckBox();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.edVND = this.Factory.CreateRibbonEditBox();
            this.edPrefix = this.Factory.CreateRibbonEditBox();
            this.edSuffix = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.box4 = this.Factory.CreateRibbonBox();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnDutoan = this.Factory.CreateRibbonButton();
            this.btnEstimateSettings = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.btnRandom3m = this.Factory.CreateRibbonButton();
            this.btnRandom5m = this.Factory.CreateRibbonButton();
            this.btnThrd = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.btPickO = this.Factory.CreateRibbonButton();
            this.btChuyenDoi = this.Factory.CreateRibbonButton();
            this.Cong = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
            this.group2.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            this.group3.SuspendLayout();
            this.box1.SuspendLayout();
            this.group4.SuspendLayout();
            this.box4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "MyTools";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.box2);
            this.group1.Items.Add(this.box3);
            this.group1.Items.Add(this.btnDutoan);
            this.group1.Items.Add(this.btnEstimateSettings);
            this.group1.Label = "Cập nhật lại đơn giá";
            this.group1.Name = "group1";
            // 
            // box2
            // 
            this.box2.Items.Add(this.checkBox1);
            this.box2.Items.Add(this.comboBox1);
            this.box2.Name = "box2";
            // 
            // checkBox1
            // 
            this.checkBox1.Checked = true;
            this.checkBox1.Label = "Active Sheets";
            this.checkBox1.Name = "checkBox1";
            // 
            // comboBox1
            // 
            this.comboBox1.Label = "comboBox1";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.ShowLabel = false;
            this.comboBox1.Text = null;
            this.comboBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBox1_TextChanged);
            // 
            // box3
            // 
            this.box3.Items.Add(this.EdCottinhtoan);
            this.box3.Items.Add(this.button1);
            this.box3.Name = "box3";
            // 
            // EdCottinhtoan
            // 
            this.EdCottinhtoan.Label = "Cột tính toán";
            this.EdCottinhtoan.Name = "EdCottinhtoan";
            this.EdCottinhtoan.Text = null;
            // 
            // group2
            // 
            ribbonDialogLauncherImpl1.ScreenTip = "Cài đặt hố đào";
            this.group2.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group2.Items.Add(this.buttonGroup1);
            this.group2.Items.Add(this.buttonGroup2);
            this.group2.Items.Add(this.btnThrd);
            this.group2.Items.Add(this.chkRownew);
            this.group2.Items.Add(this.button6);
            this.group2.Label = "Hố đào";
            this.group2.Name = "group2";
            this.group2.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group2_DialogLauncherClick);
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.button3);
            this.buttonGroup1.Items.Add(this.button4);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.btnRandom3m);
            this.buttonGroup2.Items.Add(this.btnRandom5m);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // chkRownew
            // 
            this.chkRownew.Label = "Chèn dòng mới";
            this.chkRownew.Name = "chkRownew";
            // 
            // group3
            // 
            this.group3.Items.Add(this.box1);
            this.group3.Items.Add(this.edPrefix);
            this.group3.Items.Add(this.edSuffix);
            this.group3.Items.Add(this.separator1);
            this.group3.Items.Add(this.btChuyenDoi);
            this.group3.Label = "Chuyển đổi VNĐ";
            this.group3.Name = "group3";
            // 
            // box1
            // 
            this.box1.Items.Add(this.edVND);
            this.box1.Items.Add(this.btPickO);
            this.box1.Name = "box1";
            // 
            // edVND
            // 
            this.edVND.Label = "editBox1";
            this.edVND.Name = "edVND";
            this.edVND.ShowLabel = false;
            this.edVND.Text = null;
            // 
            // edPrefix
            // 
            this.edPrefix.Label = "Prefix";
            this.edPrefix.Name = "edPrefix";
            this.edPrefix.Text = null;
            // 
            // edSuffix
            // 
            this.edSuffix.Label = "Suffix";
            this.edSuffix.Name = "edSuffix";
            this.edSuffix.Text = null;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.box4);
            this.group4.Items.Add(this.button2);
            this.group4.Items.Add(this.button5);
            this.group4.Label = "group4";
            this.group4.Name = "group4";
            // 
            // box4
            // 
            this.box4.Items.Add(this.editBox1);
            this.box4.Items.Add(this.Cong);
            this.box4.Name = "box4";
            // 
            // editBox1
            // 
            this.editBox1.Label = "editBox1";
            this.editBox1.Name = "editBox1";
            this.editBox1.ShowLabel = false;
            this.editBox1.Text = null;
            // 
            // button1
            // 
            this.button1.Label = "Cập nhật";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btnDutoan
            // 
            this.btnDutoan.Label = "Tinh toán dự toán";
            this.btnDutoan.Name = "btnDutoan";
            this.btnDutoan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDutoan_Click);
            // 
            // btnEstimateSettings
            // 
            this.btnEstimateSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEstimateSettings.Label = "Thiết lập\nChung";
            this.btnEstimateSettings.Name = "btnEstimateSettings";
            this.btnEstimateSettings.ShowImage = true;
            this.btnEstimateSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEstimateSettings_Click);
            // 
            // 
            // button3
            // 
            this.button3.Label = "Hố đào 3m";
            this.button3.Name = "button3";
            this.button3.ScreenTip = "TInh toan ho dao 3m";
            this.button3.SuperTip = "noi dung can mo ta";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Label = "Hố đào 5m";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // btnRandom3m
            // 
            this.btnRandom3m.Label = "Random3m";
            this.btnRandom3m.Name = "btnRandom3m";
            this.btnRandom3m.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRandom3m_Click);
            // 
            // btnRandom5m
            // 
            this.btnRandom5m.Label = "Random5m";
            this.btnRandom5m.Name = "btnRandom5m";
            this.btnRandom5m.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRandom5m_Click);
            // 
            // btnThrd
            // 
            this.btnThrd.Label = "RandomTH";
            this.btnThrd.Name = "btnThrd";
            this.btnThrd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnThrd_Click);
            // 
            // button6
            // 
            this.button6.Label = "Cài đặt";
            this.button6.Name = "button6";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // btPickO
            // 
            this.btPickO.Label = "Pick ô";
            this.btPickO.Name = "btPickO";
            this.btPickO.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btPickO_Click);
            // 
            // btChuyenDoi
            // 
            this.btChuyenDoi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btChuyenDoi.Image = global::ExcelAddIn1.Properties.Resources._9104252_reload_refresh_sync_repeat_icon;
            this.btChuyenDoi.Label = "Chuyển đổi";
            this.btChuyenDoi.Name = "btChuyenDoi";
            this.btChuyenDoi.ShowImage = true;
            this.btChuyenDoi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btChuyenDoi_Click);
            // 
            // Cong
            // 
            this.Cong.Label = "Cộng";
            this.Cong.Name = "Cong";
            this.Cong.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Cong_Click);
            // 
            // button2
            // 
            this.button2.Label = "GroupRow";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click_1);
            // 
            // button5
            // 
            this.button5.Label = "GroupCol";
            this.button5.Name = "button5";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EdCottinhtoan;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edVND;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btPickO;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edPrefix;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edSuffix;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btChuyenDoi;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Cong;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDutoan;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEstimateSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRandom3m;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRandom5m;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThrd;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkRownew;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
