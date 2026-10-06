using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelAddIn1.Core;

namespace ExcelAddIn1.LicenseTool
{
    public sealed class FrmLicenseTool : Form
    {
        private readonly TextBox txtMachineId;
        private readonly DateTimePicker dtpExpiry;
        private readonly TextBox txtLicenseKey;
        private readonly Label lblStatus;

        public FrmLicenseTool()
        {
            Text = "TTBMVN License Tool";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(680, 360);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                ColumnCount = 1,
                RowCount = 5
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            var title = new Label
            {
                Text = "Sinh key offline theo Machine ID",
                Dock = DockStyle.Fill,
                AutoSize = true,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 12)
            };
            root.Controls.Add(title, 0, 0);

            var inputTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2
            };
            inputTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            inputTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.Controls.Add(inputTable, 0, 1);

            inputTable.Controls.Add(CreateLabel("Machine ID"), 0, 0);
            txtMachineId = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F),
                CharacterCasing = CharacterCasing.Upper
            };
            inputTable.Controls.Add(txtMachineId, 1, 0);

            inputTable.Controls.Add(CreateLabel("Han su dung"), 0, 1);
            dtpExpiry = new DateTimePicker
            {
                Dock = DockStyle.Left,
                Width = 150,
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd",
                Value = DateTime.Today.AddYears(1)
            };
            inputTable.Controls.Add(dtpExpiry, 1, 1);

            var outputGroup = new GroupBox
            {
                Text = "License key",
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                Margin = new Padding(0, 14, 0, 10)
            };
            root.Controls.Add(outputGroup, 0, 2);

            txtLicenseKey = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 11F, FontStyle.Bold),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical
            };
            outputGroup.Controls.Add(txtLicenseKey);

            lblStatus = new Label
            {
                Text = "Nhap Machine ID va chon han su dung, sau do bam Tao key.",
                Dock = DockStyle.Fill,
                AutoSize = true,
                ForeColor = Color.FromArgb(80, 80, 80),
                Margin = new Padding(0, 0, 0, 8)
            };
            root.Controls.Add(lblStatus, 0, 3);

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true
            };
            root.Controls.Add(buttons, 0, 4);

            var btnClose = new Button { Text = "Dong", Width = 90, Height = 32 };
            btnClose.Click += (sender, args) => Close();
            buttons.Controls.Add(btnClose);

            var btnCopy = new Button { Text = "Copy key", Width = 100, Height = 32 };
            btnCopy.Click += btnCopy_Click;
            buttons.Controls.Add(btnCopy);

            var btnGenerate = new Button { Text = "Tao key", Width = 100, Height = 32 };
            btnGenerate.Click += btnGenerate_Click;
            buttons.Controls.Add(btnGenerate);

            var btnClear = new Button { Text = "Xoa", Width = 80, Height = 32 };
            btnClear.Click += btnClear_Click;
            buttons.Controls.Add(btnClear);
        }

        private static Label CreateLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 8, 6)
            };
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                string machineId = txtMachineId.Text.Trim();
                if (string.IsNullOrWhiteSpace(ProductKeyCodec.NormalizeMachineId(machineId)))
                {
                    ShowStatus("Machine ID khong hop le.", isError: true);
                    txtMachineId.Focus();
                    return;
                }

                DateTime expiry = dtpExpiry.Value.Date;
                byte[] privateKey = LicensePrivateKeyStore.Load();
                string key = ProductKeyCodec.Generate(machineId, expiry, privateKey);
                if (!ProductKeyCodec.TryValidate(key, machineId, out DateTime validatedExpiry))
                    throw new InvalidOperationException("Key sinh ra khong qua duoc buoc tu kiem tra.");

                txtLicenseKey.Text = key;
                txtLicenseKey.SelectAll();
                ShowStatus("Da tao key. Han su dung: " + validatedExpiry.ToString("yyyy-MM-dd"), isError: false);
            }
            catch (Exception ex)
            {
                ShowStatus(ex.Message, isError: true);
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtLicenseKey.Text))
            {
                ShowStatus("Chua co key de copy.", isError: true);
                return;
            }

            Clipboard.SetText(txtLicenseKey.Text.Trim());
            ShowStatus("Da copy key.", isError: false);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtMachineId.Clear();
            txtLicenseKey.Clear();
            dtpExpiry.Value = DateTime.Today.AddYears(1);
            ShowStatus("Da xoa du lieu.", isError: false);
            txtMachineId.Focus();
        }

        private void ShowStatus(string message, bool isError)
        {
            lblStatus.ForeColor = isError ? Color.FromArgb(180, 30, 30) : Color.FromArgb(20, 84, 45);
            lblStatus.Text = message;
        }
    }
}
