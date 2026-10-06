using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using ExcelAddIn1.Funtion;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    public sealed class FrmSupport : Form
    {
        private readonly Excel.Workbook workbook;
        private readonly Label lblProductValue;
        private readonly Label lblVersionValue;
        private readonly Label lblMachineIdValue;
        private readonly Label lblLicenseStatusValue;
        private readonly Label lblExpiryValue;
        private readonly TextBox txtOfflineKey;
        private readonly TextBox txtOnlineKey;

        public FrmSupport()
            : this(null)
        {
        }

        public FrmSupport(Excel.Workbook workbook)
        {
            this.workbook = workbook;
            Text = "Ho tro va kich hoat";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(700, 600);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(14),
                ColumnCount = 1,
                RowCount = 5
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            Controls.Add(root);

            GroupBox grpApp = CreateGroup("Thong tin app");
            TableLayoutPanel appTable = CreateInfoTable();
            lblProductValue = AddInfoRow(appTable, "San pham", AppInfo.ProductName);
            lblVersionValue = AddInfoRow(appTable, "Version", AppInfo.Version);
            lblMachineIdValue = AddInfoRow(appTable, "Machine ID", LicenseManager.GetMachineId());
            grpApp.Controls.Add(appTable);
            root.Controls.Add(grpApp, 0, 0);

            GroupBox grpLicense = CreateGroup("Trang thai kich hoat");
            TableLayoutPanel licenseTable = CreateInfoTable();
            lblLicenseStatusValue = AddInfoRow(licenseTable, "Trang thai", string.Empty);
            lblExpiryValue = AddInfoRow(licenseTable, "Han su dung", string.Empty);
            grpLicense.Controls.Add(licenseTable);
            root.Controls.Add(grpLicense, 0, 1);

            GroupBox grpOffline = CreateGroup("Kich hoat offline");
            var offlinePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                ColumnCount = 3,
                RowCount = 3
            };
            offlinePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105));
            offlinePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            offlinePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
            offlinePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            offlinePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            offlinePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lblOfflineNote = new Label
            {
                Text = "Dan key offline duoc cap theo Machine ID cua may nay.",
                Dock = DockStyle.Fill,
                AutoSize = true,
                ForeColor = Color.FromArgb(90, 90, 90),
                Margin = new Padding(3, 0, 3, 6)
            };
            offlinePanel.Controls.Add(lblOfflineNote, 0, 0);
            offlinePanel.SetColumnSpan(lblOfflineNote, 3);

            offlinePanel.Controls.Add(CreateLabel("Nhap key"), 0, 1);
            txtOfflineKey = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                Height = 24
            };
            offlinePanel.Controls.Add(txtOfflineKey, 1, 1);
            var btnActivateOffline = new Button { Text = "Kich hoat", Dock = DockStyle.Fill };
            btnActivateOffline.Click += btnActivateOffline_Click;
            offlinePanel.Controls.Add(btnActivateOffline, 2, 1);

            var btnCopyMachineId = new Button { Text = "Copy Machine ID", Width = 140, Height = 28 };
            btnCopyMachineId.Click += btnCopyMachineId_Click;
            offlinePanel.Controls.Add(btnCopyMachineId, 1, 2);
            grpOffline.Controls.Add(offlinePanel);
            root.Controls.Add(grpOffline, 0, 2);

            GroupBox grpOnline = CreateGroup("Kich hoat online");
            var onlinePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                ColumnCount = 3,
                RowCount = 2
            };
            onlinePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            onlinePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            onlinePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
            onlinePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            onlinePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            onlinePanel.Controls.Add(CreateLabel("Online key"), 0, 0);
            txtOnlineKey = new TextBox { Dock = DockStyle.Fill, Enabled = false };
            onlinePanel.Controls.Add(txtOnlineKey, 1, 0);
            var btnActivateOnline = new Button { Text = "Tam tat", Dock = DockStyle.Fill, Enabled = false };
            onlinePanel.Controls.Add(btnActivateOnline, 2, 0);
            var lblOnlineNote = new Label
            {
                Text = "Kich hoat online dang tam tat. Hien tai chi dung kich hoat offline.",
                Dock = DockStyle.Fill,
                AutoSize = true,
                ForeColor = Color.FromArgb(110, 110, 110)
            };
            onlinePanel.Controls.Add(lblOnlineNote, 1, 1);
            grpOnline.Controls.Add(onlinePanel);
            root.Controls.Add(grpOnline, 0, 3);

            var bottomPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = false,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0)
            };
            var btnClose = new Button { Text = "Dong", Width = 90, Height = 30 };
            btnClose.Click += (sender, args) => Close();
            var btnCreateSupport = new Button { Text = "Tao goi chan doan", Width = 150, Height = 30 };
            btnCreateSupport.Click += btnCreateSupport_Click;
            var btnOpenLog = new Button { Text = "Mo thu muc log", Width = 130, Height = 30 };
            btnOpenLog.Click += btnOpenLog_Click;
            bottomPanel.Controls.Add(btnClose);
            bottomPanel.Controls.Add(btnCreateSupport);
            bottomPanel.Controls.Add(btnOpenLog);
            root.Controls.Add(bottomPanel, 0, 4);

            RefreshLicenseStatus();
        }

        private static GroupBox CreateGroup(string text)
        {
            return new GroupBox
            {
                Text = text,
                Dock = DockStyle.Fill,
                AutoSize = false,
                Padding = new Padding(8)
            };
        }

        private static TableLayoutPanel CreateInfoTable()
        {
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Padding = new Padding(8)
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return table;
        }

        private static Label AddInfoRow(TableLayoutPanel table, string label, string value)
        {
            int row = table.RowCount++;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(CreateLabel(label), 0, row);

            var valueLabel = new Label
            {
                Text = value,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            table.Controls.Add(valueLabel, 1, row);
            return valueLabel;
        }

        private static Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(3, 6, 3, 6)
            };
        }

        private void RefreshLicenseStatus()
        {
            LicenseStatus status = LicenseManager.GetStatus();
            lblProductValue.Text = AppInfo.ProductName;
            lblVersionValue.Text = AppInfo.Version;
            lblMachineIdValue.Text = status.MachineId;
            lblLicenseStatusValue.Text = status.Message;
            lblLicenseStatusValue.ForeColor = status.IsExpired ? Color.FromArgb(180, 30, 30) : Color.FromArgb(20, 84, 45);

            if (status.ExpiryDate.HasValue)
                lblExpiryValue.Text = status.ExpiryDate.Value.ToString("yyyy-MM-dd");
            else if (status.IsTrial)
                lblExpiryValue.Text = "Dung thu con " + status.TrialDaysRemaining + " ngay";
            else
                lblExpiryValue.Text = "Khong co";
        }

        private void btnActivateOffline_Click(object sender, EventArgs e)
        {
            try
            {
                if (LicenseManager.TryActivate(txtOfflineKey.Text, out string message))
                {
                    RefreshLicenseStatus();
                    MessageBox.Show(message, "Kich hoat offline", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MessageBox.Show(message, "Kich hoat offline", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Kich hoat offline", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnCopyMachineId_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(lblMachineIdValue.Text);
            MessageBox.Show("Da copy Machine ID.", "Ho tro", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnCreateSupport_Click(object sender, EventArgs e)
        {
            try
            {
                string folder = workbook == null
                    ? SupportPackage.Create()
                    : SupportPackage.Create(workbook);
                OpenFolder(folder);
                MessageBox.Show("Da tao goi chan doan tai:" + Environment.NewLine + folder, "Ho tro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Khong tao duoc goi chan doan", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnOpenLog_Click(object sender, EventArgs e)
        {
            try
            {
                System.IO.Directory.CreateDirectory(AppPaths.LogDirectory);
                OpenFolder(AppPaths.LogDirectory);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Khong mo duoc thu muc log", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static void OpenFolder(string folder)
        {
            Process.Start(new ProcessStartInfo("explorer.exe", "\"" + folder + "\"") { UseShellExecute = true });
        }
    }
}
