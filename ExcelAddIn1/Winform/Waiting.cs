using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn1.Winform
{
    public partial class Waiting : Form
    {
        private PictureBox picLoading;
        private Label lblCaption;
        private Label lblDescription;
        public Waiting(string caption = "Vui lòng chờ!", string description = "Đang xử lý dữ liệu...")
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;
            this.Opacity = 0.95;
            this.Size = new Size(320, 120);
            this.ShowInTaskbar = false;
            this.TopMost = true;

            // Ảnh GIF loading
            picLoading = new PictureBox
            {
                Size = new Size(36, 36),
                Location = new Point((this.Width - 36) / 2 - 80, (this.Height - 36) / 2 - 10),
                Image = Properties.Resources.Refesh_1, // GIF nền trong suốt
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.Transparent
            };

            // Dòng tiêu đề
            lblCaption = new Label
            {
                Text = caption,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Size = new Size(200, 28),
                Location = new Point(picLoading.Right + 10, picLoading.Top)
            };

            // Dòng mô tả
            lblDescription = new Label
            {
                Text = description,
                Font = new Font("Segoe UI", 10, FontStyle.Regular),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Size = new Size(240, 24),
                Location = new Point(picLoading.Right + 10, lblCaption.Bottom - 5),
                ForeColor = Color.DimGray
            };

            this.Controls.Add(picLoading);
            this.Controls.Add(lblCaption);
            this.Controls.Add(lblDescription);

            // Bo góc form
            this.Region = System.Drawing.Region.FromHrgn(
                CreateRoundRectRgn(0, 0, this.Width, this.Height, 12, 12));
        }
        public void UpdateDescription(string text)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => lblDescription.Text = text));
            }
            else
            {
                lblDescription.Text = text;
            }
        }


        // API bo góc
        [System.Runtime.InteropServices.DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);
    }
}
