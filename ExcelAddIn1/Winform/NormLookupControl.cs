using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace ExcelAddIn1.Winform
{
    public sealed class NormLookupControl : UserControl
    {
        private readonly NormSearchIndex searchIndex;
        private readonly TextBox searchTextBox;
        private readonly ComboBox environmentComboBox;
        private readonly ComboBox depthComboBox;
        private readonly ComboBox resourceComboBox;
        private readonly ComboBox variantComboBox;
        private readonly DataGridView resultGrid;
        private readonly Label titleLabel;
        private readonly Label metaLabel;
        private readonly Label sourceLabel;
        private readonly Label conditionLabel;
        private readonly DataGridView rateGrid;
        private readonly Label statusLabel;

        public NormLookupControl(NormSearchIndex searchIndex)
        {
            this.searchIndex = searchIndex ?? throw new ArgumentNullException(nameof(searchIndex));
            Dock = DockStyle.Fill;
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = SystemColors.Control;

            var filters = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 78,
                Padding = new Padding(10, 8, 10, 6),
                ColumnCount = 5,
                RowCount = 2
            };
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 17));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            filters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 15));
            filters.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            filters.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

            filters.Controls.Add(CreateFilterLabel("Mã hoặc nội dung"), 0, 0);
            filters.Controls.Add(CreateFilterLabel("Môi trường"), 1, 0);
            filters.Controls.Add(CreateFilterLabel("Độ sâu (m)"), 2, 0);
            filters.Controls.Add(CreateFilterLabel("Nguồn lực"), 3, 0);
            filters.Controls.Add(CreateFilterLabel("Biến thể"), 4, 0);

            searchTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = SystemColors.Window
            };
            environmentComboBox = CreateFilterComboBox();
            depthComboBox = CreateFilterComboBox();
            resourceComboBox = CreateFilterComboBox();
            variantComboBox = CreateFilterComboBox();
            variantComboBox.DropDownStyle = ComboBoxStyle.DropDown;
            filters.Controls.Add(searchTextBox, 0, 1);
            filters.Controls.Add(environmentComboBox, 1, 1);
            filters.Controls.Add(depthComboBox, 2, 1);
            filters.Controls.Add(resourceComboBox, 3, 1);
            filters.Controls.Add(variantComboBox, 4, 1);

            resultGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle
            };
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Mã",
                Width = 118
            });
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Tên định mức",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Đơn vị",
                Width = 92
            });
            resultGrid.SelectionChanged += (sender, args) => ShowSelectedResult();

            var detailPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 6, 10, 6)
            };
            titleLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 48,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                AutoEllipsis = true
            };
            metaLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 42,
                ForeColor = SystemColors.ControlText,
                AutoEllipsis = true
            };
            sourceLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 52,
                ForeColor = SystemColors.GrayText,
                AutoEllipsis = true
            };
            conditionLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 64,
                ForeColor = SystemColors.ControlText,
                AutoEllipsis = true
            };
            rateGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle
            };
            detailPanel.Controls.Add(rateGrid);
            detailPanel.Controls.Add(conditionLabel);
            detailPanel.Controls.Add(sourceLabel);
            detailPanel.Controls.Add(metaLabel);
            detailPanel.Controls.Add(titleLabel);

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Size = new Size(800, 500),
                SplitterDistance = 390,
                Panel1MinSize = 300,
                Panel2MinSize = 360
            };
            split.Panel1.Padding = new Padding(10, 4, 4, 6);
            split.Panel2.Padding = new Padding(4, 4, 10, 6);
            split.Panel1.Controls.Add(resultGrid);
            split.Panel2.Controls.Add(detailPanel);

            statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                Padding = new Padding(10, 6, 10, 4),
                ForeColor = SystemColors.GrayText
            };

            Controls.Add(split);
            Controls.Add(filters);
            Controls.Add(statusLabel);

            PopulateFilters();
            searchTextBox.TextChanged += (sender, args) => PerformSearch();
            environmentComboBox.SelectedIndexChanged += (sender, args) => PerformSearch();
            depthComboBox.SelectedIndexChanged += (sender, args) => PerformSearch();
            resourceComboBox.SelectedIndexChanged += (sender, args) => PerformSearch();
            variantComboBox.TextChanged += (sender, args) => PerformSearch();
            PerformSearch();
        }

        private void PopulateFilters()
        {
            environmentComboBox.Items.Add(new FilterItem<NormEnvironment?>("Tất cả", null));
            environmentComboBox.Items.Add(new FilterItem<NormEnvironment?>("Khảo sát", NormEnvironment.Survey));
            environmentComboBox.Items.Add(new FilterItem<NormEnvironment?>("Trên cạn", NormEnvironment.Land));
            environmentComboBox.Items.Add(new FilterItem<NormEnvironment?>("Dưới nước", NormEnvironment.InlandWater));
            environmentComboBox.Items.Add(new FilterItem<NormEnvironment?>("Dưới biển", NormEnvironment.Sea));
            environmentComboBox.SelectedIndex = 0;

            depthComboBox.Items.Add(new FilterItem<decimal?>("Tất cả", null));
            IReadOnlyList<NormSearchResult> all = searchIndex.Search(
                new NormSearchQuery(string.Empty, maximumResults: 500));
            foreach (decimal depth in all.SelectMany(result => result.DepthMeters).Distinct().OrderBy(value => value))
            {
                depthComboBox.Items.Add(new FilterItem<decimal?>(
                    depth.ToString("0.##", CultureInfo.CurrentCulture),
                    depth));
            }
            depthComboBox.SelectedIndex = 0;

            resourceComboBox.Items.Add(new FilterItem<NormResourceKind?>("Tất cả", null));
            resourceComboBox.Items.Add(new FilterItem<NormResourceKind?>("Vật liệu", NormResourceKind.Material));
            resourceComboBox.Items.Add(new FilterItem<NormResourceKind?>("Nhân công", NormResourceKind.Labor));
            resourceComboBox.Items.Add(new FilterItem<NormResourceKind?>("Máy", NormResourceKind.Machine));
            resourceComboBox.SelectedIndex = 0;

            variantComboBox.Items.Add(string.Empty);
            foreach (string variant in all
                .Where(result => result.Definition != null)
                .SelectMany(result => result.Definition.Variants)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal))
            {
                variantComboBox.Items.Add(variant);
            }
            variantComboBox.SelectedIndex = 0;
        }

        private void PerformSearch()
        {
            if (environmentComboBox.SelectedItem == null ||
                depthComboBox.SelectedItem == null ||
                resourceComboBox.SelectedItem == null)
            {
                return;
            }

            try
            {
                var environment = ((FilterItem<NormEnvironment?>)environmentComboBox.SelectedItem).Value;
                var depth = ((FilterItem<decimal?>)depthComboBox.SelectedItem).Value;
                var resource = ((FilterItem<NormResourceKind?>)resourceComboBox.SelectedItem).Value;
                IReadOnlyList<NormSearchResult> results = searchIndex.Search(
                    new NormSearchQuery(
                        searchTextBox.Text,
                        environment,
                        depth,
                        variantComboBox.Text,
                        resource,
                        500));

                resultGrid.Rows.Clear();
                foreach (NormSearchResult result in results)
                {
                    int rowIndex = resultGrid.Rows.Add(
                        result.Key,
                        result.Title,
                        result.WorkUnit);
                    resultGrid.Rows[rowIndex].Tag = result;
                }
                statusLabel.ForeColor = SystemColors.GrayText;
                statusLabel.Text = results.Count.ToString(CultureInfo.CurrentCulture) +
                    " định mức trong package đang dùng";
                if (resultGrid.Rows.Count > 0)
                    resultGrid.Rows[0].Selected = true;
                else
                    ClearDetails();
            }
            catch (Exception ex)
            {
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = ex.Message;
            }
        }

        private void ShowSelectedResult()
        {
            if (resultGrid.SelectedRows.Count != 1 ||
                !(resultGrid.SelectedRows[0].Tag is NormSearchResult result))
            {
                ClearDetails();
                return;
            }

            titleLabel.Text = result.Key + " - " + result.Title;
            metaLabel.Text = "Đơn vị: " + result.WorkUnit +
                "    Package: " + result.PackageId + " v" + result.PackageVersion;
            sourceLabel.Text = "Căn cứ: " + result.Source.DocumentId +
                ", trang " + result.Source.PageFrom.ToString(CultureInfo.CurrentCulture) +
                (result.Source.PageTo == result.Source.PageFrom
                    ? string.Empty
                    : "-" + result.Source.PageTo.ToString(CultureInfo.CurrentCulture)) +
                ", " + result.Source.Section;

            rateGrid.Columns.Clear();
            rateGrid.Rows.Clear();
            if (result.Definition == null)
            {
                conditionLabel.Text = "Package này chưa có bảng hao phí chi tiết trong dữ liệu số hóa.";
                return;
            }

            rateGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Loại",
                Width = 66
            });
            rateGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Mã nguồn lực",
                Width = 118
            });
            rateGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Đơn vị",
                Width = 68
            });
            foreach (string variant in result.Definition.Variants)
            {
                rateGrid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    HeaderText = variant,
                    Width = 55
                });
            }
            foreach (NormResourceRate rate in result.Definition.Rates)
            {
                var values = new List<object>
                {
                    GetResourceLabel(rate.Kind),
                    rate.ResourceCode,
                    rate.Unit
                };
                values.AddRange(rate.Quantities.Select(value =>
                    (object)value.ToString("0.####", CultureInfo.CurrentCulture)));
                rateGrid.Rows.Add(values.ToArray());
            }

            string adjustments = result.Definition.Adjustments.Count == 0
                ? "Không có điều chỉnh."
                : "Điều chỉnh: " + string.Join(", ", result.Definition.Adjustments.Select(item => item.Condition));
            string constraints = result.Definition.Constraints.Count == 0
                ? "Không có ràng buộc."
                : "Ràng buộc: " + string.Join(", ", result.Definition.Constraints.Select(item =>
                    item.Operation + ":" + item.Argument1));
            conditionLabel.Text = adjustments + "  " + constraints;
        }

        private void ClearDetails()
        {
            titleLabel.Text = string.Empty;
            metaLabel.Text = string.Empty;
            sourceLabel.Text = string.Empty;
            conditionLabel.Text = string.Empty;
            rateGrid.Columns.Clear();
            rateGrid.Rows.Clear();
        }

        private static Label CreateFilterLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static ComboBox CreateFilterComboBox()
        {
            return new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
        }

        private static string GetResourceLabel(NormResourceKind kind)
        {
            switch (kind)
            {
                case NormResourceKind.Material: return "Vật liệu";
                case NormResourceKind.Labor: return "Nhân công";
                case NormResourceKind.Machine: return "Máy";
                default: return kind.ToString();
            }
        }

        private sealed class FilterItem<T>
        {
            public FilterItem(string text, T value)
            {
                Text = text;
                Value = value;
            }

            public string Text { get; }
            public T Value { get; }

            public override string ToString()
            {
                return Text;
            }
        }
    }
}
