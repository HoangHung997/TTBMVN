using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelAddIn1.Winform
{
    internal abstract class EstimateActionPane : UserControl
    {
        protected static readonly Color Green = Color.FromArgb(0, 137, 70);
        protected static readonly Color Blue = Color.FromArgb(45, 112, 229);
        protected static readonly Color Red = Color.FromArgb(220, 53, 69);
        protected readonly TableLayoutPanel Content;
        protected readonly Label Status;

        protected EstimateActionPane(string title, EstimateUiIconKind icon, Action back)
        {
            Dock = DockStyle.Fill; BackColor = Color.White; AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 9f);
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            Controls.Add(scroll);
            Content = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1,
                Padding = new Padding(14, 10, 14, 14), BackColor = Color.White };
            Content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            scroll.Controls.Add(Content);
            var header = new TableLayoutPanel { Dock = DockStyle.Top, Height = 56, ColumnCount = 2 };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 36));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.Controls.Add(new PictureBox { Dock = DockStyle.Fill, Image = EstimateUiIcons.Create(icon, 28, Green),
                SizeMode = PictureBoxSizeMode.CenterImage }, 0, 0);
            header.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Trợ lý Dự toán / " + title,
                Font = new Font("Segoe UI", 12f, FontStyle.Bold), ForeColor = Color.FromArgb(0, 103, 55),
                TextAlign = ContentAlignment.MiddleLeft }, 1, 0);
            Add(header);
            Add(Command("Tổng quan", EstimateUiIconKind.Clipboard, back, false));
            Status = new Label { AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(10),
                BackColor = Color.FromArgb(237, 244, 255), ForeColor = Blue, MinimumSize = new Size(0, 45) };
        }

        protected void Add(Control control)
        {
            control.Dock = DockStyle.Top;
            control.Margin = new Padding(0, 4, 0, 7);
            Content.Controls.Add(control, 0, Content.RowCount++);
        }

        protected Label Section(string text)
        {
            var label = new Label { Text = text, AutoSize = true, Font = new Font(Font, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 43, 54), Padding = new Padding(0, 8, 0, 0) };
            Add(label); return label;
        }

        protected static Button Command(string text, EstimateUiIconKind icon, Action action, bool primary = true)
        {
            var button = new Button { Text = text, Height = 36, FlatStyle = FlatStyle.Flat,
                BackColor = primary ? Green : Color.White, ForeColor = primary ? Color.White : Blue,
                TextImageRelation = TextImageRelation.ImageBeforeText, ImageAlign = ContentAlignment.MiddleLeft,
                Image = EstimateUiIcons.Create(icon, 18, primary ? Color.White : Blue),
                UseVisualStyleBackColor = false, AutoEllipsis = true };
            button.FlatAppearance.BorderColor = Color.FromArgb(222, 228, 223);
            button.Click += (s, e) => action?.Invoke();
            return button;
        }

        protected void SetStatus(string message, bool error = false)
        {
            Status.Text = message;
            Status.ForeColor = error ? Red : Blue;
        }

        protected void Run(Action action)
        {
            try { action(); }
            catch (Exception ex)
            {
                Funtion.RuntimeLogger.Log(ex, "Estimate V2 action pane");
                SetStatus(ex.Message, true);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) DisposeImages(this);
            base.Dispose(disposing);
        }
        private static void DisposeImages(Control root)
        {
            foreach (Control control in root.Controls)
            {
                var button = control as Button;
                if (button?.Image != null) { button.Image.Dispose(); button.Image = null; }
                var picture = control as PictureBox;
                if (picture?.Image != null) { picture.Image.Dispose(); picture.Image = null; }
                DisposeImages(control);
            }
        }
    }
}
