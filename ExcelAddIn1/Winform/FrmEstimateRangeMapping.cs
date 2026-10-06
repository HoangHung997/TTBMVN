using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal sealed class FrmEstimateRangeMapping : Form
    {
        private readonly Dictionary<string, ComboBox> inputs = new Dictionary<string, ComboBox>();
        private readonly CheckBox headerInput;
        private readonly ComboBox environmentInput;
        private readonly ComboBox audienceInput;

        internal FrmEstimateRangeMapping(Excel.Range range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));
            Text = "Ánh xạ cột Phụ lục dự toán";
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            ClientSize = new Size(650, 555);
            AutoScaleMode = AutoScaleMode.Dpi;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 4
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));

            Excel.Worksheet worksheet = null;
            string worksheetName;
            try
            {
                worksheet = range.Worksheet;
                worksheetName = worksheet.Name;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
            root.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Text = worksheetName + "!" + range.Address[true, true, Excel.XlReferenceStyle.xlA1] +
                    "\nChọn cột nguồn và các cột app sẽ ghi công thức kết quả."
            }, 0, 0);

            var defaults = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            headerInput = new CheckBox
            {
                Text = "Dòng đầu là tiêu đề",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(0, 10, 24, 0)
            };
            environmentInput = DropDown(120);
            environmentInput.Items.Add(new Option<EstimateWorkEnvironment>("Cạn", EstimateWorkEnvironment.Land));
            environmentInput.Items.Add(new Option<EstimateWorkEnvironment>("Nước", EstimateWorkEnvironment.Water));
            environmentInput.SelectedIndex = 0;
            audienceInput = DropDown(190);
            audienceInput.Items.Add(new Option<MachineRateAudience>(
                "Hưởng lương NSNN", MachineRateAudience.StateBudgetSalary));
            audienceInput.Items.Add(new Option<MachineRateAudience>(
                "Không hưởng lương NSNN", MachineRateAudience.NonStateSalary));
            audienceInput.SelectedIndex = 1;
            defaults.Controls.Add(headerInput);
            defaults.Controls.Add(InlineLabel("Mặc định:"));
            defaults.Controls.Add(environmentInput);
            defaults.Controls.Add(audienceInput);
            root.Controls.Add(defaults, 0, 1);

            IReadOnlyList<ColumnOption> columns = ReadColumns(range);
            var mapping = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                ColumnCount = 2,
                RowCount = 11,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            };
            mapping.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 205));
            mapping.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            AddMap(mapping, 0, "code", "Số hiệu/Mã công tác", columns, false, "SO HIEU", "MA HIEU", "MA CONG TAC");
            AddMap(mapping, 1, "description", "Nội dung công tác *", columns, true, "TEN CONG TAC", "NOI DUNG", "TEN HANG MUC");
            AddMap(mapping, 2, "unit", "Đơn vị", columns, false, "DVT", "DON VI");
            AddMap(mapping, 3, "quantity", "Khối lượng *", columns, true, "KHOI LUONG", "KL");
            AddMap(mapping, 4, "accepted", "Khối lượng nghiệm thu", columns, false, "KL NGHIEM THU", "KHOI LUONG NGHIEM THU");
            AddMap(mapping, 5, "rateMaterial", "Đơn giá vật liệu", columns, false, "DG VL", "DON GIA VAT LIEU");
            AddMap(mapping, 6, "rateLabor", "Đơn giá nhân công", columns, false, "DG NC", "DON GIA NHAN CONG");
            AddMap(mapping, 7, "rateMachine", "Đơn giá máy", columns, false, "DG M", "DON GIA MAY");
            AddMap(mapping, 8, "amountMaterial", "Thành tiền vật liệu", columns, false, "TT VL", "THANH TIEN VAT LIEU");
            AddMap(mapping, 9, "amountLabor", "Thành tiền nhân công", columns, false, "TT NC", "THANH TIEN NHAN CONG");
            AddMap(mapping, 10, "amountMachine", "Thành tiền máy", columns, false, "TT M", "THANH TIEN MAY");
            root.Controls.Add(mapping, 0, 2);

            var commands = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 8, 0, 0)
            };
            var ok = new Button { Text = "Đọc vùng", DialogResult = DialogResult.OK, Width = 100, Height = 30 };
            var cancel = new Button { Text = "Hủy", DialogResult = DialogResult.Cancel, Width = 85, Height = 30 };
            ok.Click += ValidateBeforeClose;
            commands.Controls.Add(ok);
            commands.Controls.Add(cancel);
            root.Controls.Add(commands, 0, 3);
            Controls.Add(root);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        internal bool FirstRowIsHeader => headerInput.Checked;
        internal EstimateWorkEnvironment DefaultEnvironment =>
            ((Option<EstimateWorkEnvironment>)environmentInput.SelectedItem).Value;
        internal MachineRateAudience DefaultAudience =>
            ((Option<MachineRateAudience>)audienceInput.SelectedItem).Value;

        internal EstimateColumnMap BuildColumnMap()
        {
            return new EstimateColumnMap(
                Column("code"),
                Column("description"),
                Column("unit"),
                Column("quantity"),
                Column("accepted"),
                Column("rateMaterial"),
                Column("rateLabor"),
                Column("rateMachine"),
                Column("amountMaterial"),
                Column("amountLabor"),
                Column("amountMachine"));
        }

        private void ValidateBeforeClose(object sender, EventArgs e)
        {
            if (Column("description") > 0 && Column("quantity") > 0)
                return;
            MessageBox.Show(
                "Phải chọn cột Nội dung công tác và Khối lượng.",
                Text,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            DialogResult = DialogResult.None;
        }

        private int Column(string key)
        {
            return (inputs[key].SelectedItem as ColumnOption)?.ColumnNumber ?? 0;
        }

        private void AddMap(
            TableLayoutPanel panel,
            int row,
            string key,
            string title,
            IReadOnlyList<ColumnOption> columns,
            bool required,
            params string[] candidates)
        {
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            panel.Controls.Add(new Label
            {
                Text = title,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(7, 0, 0, 0)
            }, 0, row);
            ComboBox combo = DropDown(100);
            combo.Dock = DockStyle.Fill;
            combo.Margin = new Padding(6, 4, 6, 4);
            combo.Items.Add(new ColumnOption(0, "(Không dùng)"));
            foreach (ColumnOption option in columns)
                combo.Items.Add(option);
            int match = FindBestColumn(columns, candidates);
            combo.SelectedIndex = match >= 0 ? match + 1 : required && columns.Count > row ? row + 1 : 0;
            inputs.Add(key, combo);
            panel.Controls.Add(combo, 1, row);
        }

        private static int FindBestColumn(IReadOnlyList<ColumnOption> columns, IEnumerable<string> candidates)
        {
            string[] wanted = candidates.Select(Normalize).ToArray();
            for (int index = 0; index < columns.Count; index++)
            {
                string header = Normalize(columns[index].Header);
                if (wanted.Any(candidate => header == candidate || header.Contains(candidate)))
                    return index;
            }
            return -1;
        }

        private static IReadOnlyList<ColumnOption> ReadColumns(Excel.Range range)
        {
            object values = range.Value2;
            var result = new List<ColumnOption>();
            for (int index = 1; index <= range.Columns.Count; index++)
            {
                object raw = MatrixValue(values, 1, index);
                string header = (Convert.ToString(raw, CultureInfo.CurrentCulture) ?? string.Empty).Trim();
                int column = range.Column + index - 1;
                string text = ExcelColumnAddress.ToLetters(column) +
                    (header.Length == 0 ? string.Empty : " - " + header);
                result.Add(new ColumnOption(column, text, header));
            }
            return result;
        }

        private static object MatrixValue(object matrix, int row, int column)
        {
            Array array = matrix as Array;
            if (array == null)
                return row == 1 && column == 1 ? matrix : null;
            return array.GetValue(
                array.GetLowerBound(0) + row - 1,
                array.GetLowerBound(1) + column - 1);
        }

        private static string Normalize(string value)
        {
            string decomposed = (value ?? string.Empty).Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder();
            foreach (char character in decomposed)
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(character);
                if (category != UnicodeCategory.NonSpacingMark)
                    builder.Append(char.ToUpperInvariant(character));
            }
            return string.Join(" ", builder.ToString().Split(
                new[] { ' ', '\t', '\r', '\n', '-', '_', '/', '(', ')' },
                StringSplitOptions.RemoveEmptyEntries));
        }

        private static ComboBox DropDown(int width)
        {
            return new ComboBox
            {
                Width = width,
                Height = 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                IntegralHeight = false,
                MaxDropDownItems = 16
            };
        }

        private static Label InlineLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 10, 5, 0)
            };
        }

        private sealed class ColumnOption
        {
            internal ColumnOption(int columnNumber, string text, string header = "")
            {
                ColumnNumber = columnNumber;
                Text = text;
                Header = header;
            }

            internal int ColumnNumber { get; }
            internal string Text { get; }
            internal string Header { get; }
            public override string ToString() => Text;
        }

        private sealed class Option<T>
        {
            internal Option(string text, T value)
            {
                Text = text;
                Value = value;
            }

            internal string Text { get; }
            internal T Value { get; }
            public override string ToString() => Text;
        }
    }
}
