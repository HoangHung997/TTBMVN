using ExcelAddIn1.Core;
using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal sealed class EstimateNormBindingPaneView : UserControl
    {
        private static readonly Color Green = Color.FromArgb(0, 137, 70);
        private static readonly Color GreenDark = Color.FromArgb(0, 103, 55);
        private static readonly Color GreenSoft = Color.FromArgb(235, 248, 239);
        private static readonly Color Border = Color.FromArgb(222, 228, 223);
        private static readonly Color TextDark = Color.FromArgb(33, 43, 54);
        private static readonly Color TextMuted = Color.FromArgb(92, 103, 112);
        private static readonly Color Amber = Color.FromArgb(244, 166, 35);
        private static readonly Color Red = Color.FromArgb(220, 53, 69);
        private static readonly Color Blue = Color.FromArgb(45, 112, 229);

        private readonly Excel.Workbook workbook;
        private readonly Action backAction;
        private readonly Label boundValue;
        private readonly Label unboundValue;
        private readonly Label totalValue;
        private readonly TextBox searchBox;
        private readonly FlowLayoutPanel filterPanel;
        private readonly DataGridView resultGrid;
        private readonly TabControl detailTabs;
        private readonly Label statusLabel;
        private NormSearchIndex searchIndex;
        private NormEnvironment? environmentFilter;
        private IReadOnlyList<NormChoice> currentChoices = new NormChoice[0];

        internal EstimateNormBindingPaneView(
            Excel.Workbook workbook,
            Action backAction)
        {
            this.workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            this.backAction = backAction;
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;

            var root = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            Controls.Add(root);
            var content = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(14, 10, 14, 14),
                BackColor = Color.White,
                ColumnCount = 1
            };
            root.Controls.Add(content);

            content.Controls.Add(BuildHeader(), 0, content.RowCount++);

            var metrics = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 98,
                ColumnCount = 3,
                Margin = new Padding(0, 4, 0, 9)
            };
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            metrics.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.34f));
            boundValue = AddMetric(metrics, 0, "Đã gắn", EstimateUiIconKind.Check, Green);
            unboundValue = AddMetric(metrics, 1, "Chưa gắn", EstimateUiIconKind.Warning, Amber);
            totalValue = AddMetric(metrics, 2, "Định mức đã dùng", EstimateUiIconKind.Database, Blue);
            content.Controls.Add(metrics, 0, content.RowCount++);

            searchBox = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 32,
                Margin = new Padding(0, 0, 0, 7),
                Font = new Font("Segoe UI", 9f),
                BorderStyle = BorderStyle.FixedSingle
            };
            searchBox.TextChanged += (s, e) => RefreshSearch();
            content.Controls.Add(searchBox, 0, content.RowCount++);

            filterPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 38,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 7)
            };
            AddFilter("Tất cả", null, true);
            AddFilter("Trên cạn", NormEnvironment.Land, false);
            AddFilter("Dưới nước", NormEnvironment.InlandWater, false);
            AddFilter("Trên biển", NormEnvironment.Sea, false);
            content.Controls.Add(filterPanel, 0, content.RowCount++);

            resultGrid = new DataGridView
            {
                Dock = DockStyle.Top,
                Height = 215,
                Margin = new Padding(0, 0, 0, 7),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                AutoGenerateColumns = false,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ReadOnly = true,
                Font = new Font("Segoe UI", 8f)
            };
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colNormCode",
                HeaderText = "Mã định mức",
                Width = 116
            });
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colNormTitle",
                HeaderText = "Tên định mức",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 150
            });
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colNormUnit",
                HeaderText = "Đơn vị",
                Width = 62
            });
            resultGrid.SelectionChanged += (s, e) => RefreshDetail();
            content.Controls.Add(resultGrid, 0, content.RowCount++);

            var detailTitle = new Label
            {
                Text = "Chi tiết hao phí của định mức đang chọn",
                Dock = DockStyle.Top,
                Height = 31,
                Margin = new Padding(0, 0, 0, 2),
                Font = new Font("Segoe UI", 9.2f, FontStyle.Bold),
                ForeColor = TextDark,
                TextAlign = ContentAlignment.MiddleLeft
            };
            content.Controls.Add(detailTitle, 0, content.RowCount++);

            detailTabs = new TabControl
            {
                Dock = DockStyle.Top,
                Height = 190,
                Margin = new Padding(0, 0, 0, 7),
                Font = new Font("Segoe UI", 8f)
            };
            detailTabs.TabPages.Add(CreateDetailPage("Vật liệu", NormResourceKind.Material));
            detailTabs.TabPages.Add(CreateDetailPage("Nhân công", NormResourceKind.Labor));
            detailTabs.TabPages.Add(CreateDetailPage("Máy thi công", NormResourceKind.Machine));
            content.Controls.Add(detailTabs, 0, content.RowCount++);

            var buttons = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 45,
                ColumnCount = 4,
                Margin = new Padding(0, 0, 0, 7)
            };
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 31));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            Button bindCurrent = ActionButton("✓ Gắn dòng này", Green, Color.White);
            Button bindSelected = ActionButton("Gắn các dòng đã chọn", Color.White, Green);
            Button change = ActionButton("↔ Đổi định mức", Color.White, Color.FromArgb(66, 81, 94));
            Button unbind = ActionButton("Bỏ gắn", Color.White, Red);
            bindCurrent.Click += (s, e) => Bind(false);
            bindSelected.Click += (s, e) => Bind(true);
            change.Click += (s, e) => Bind(false);
            unbind.Click += (s, e) => UnbindSelected();
            buttons.Controls.Add(bindCurrent, 0, 0);
            buttons.Controls.Add(bindSelected, 1, 0);
            buttons.Controls.Add(change, 2, 0);
            buttons.Controls.Add(unbind, 3, 0);
            content.Controls.Add(buttons, 0, content.RowCount++);

            statusLabel = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                MinimumSize = new Size(0, 57),
                Padding = new Padding(10, 9, 10, 9),
                BackColor = Color.FromArgb(235, 244, 255),
                ForeColor = Color.FromArgb(35, 88, 180),
                Font = new Font("Segoe UI", 8.1f),
                Text = "Xóa nội dung ở ô định mức trên sheet không làm mất liên kết gắn định mức."
            };
            content.Controls.Add(statusLabel, 0, content.RowCount++);

            LoadSearchIndex();
            RefreshStateMetrics();
            RefreshSearch();
        }

        private void LoadSearchIndex()
        {
            try
            {
                searchIndex = WorkbookNormCatalogService.LoadPinnedSearchIndex(workbook);
                statusLabel.BackColor = Color.FromArgb(235, 244, 255);
                statusLabel.ForeColor = Color.FromArgb(35, 88, 180);
                statusLabel.Text = "Định mức được lưu ẩn theo WorkItemId. Xóa ô hiển thị trên sheet không làm mất binding.";
            }
            catch (Exception ex)
            {
                searchIndex = null;
                statusLabel.BackColor = Color.FromArgb(255, 247, 224);
                statusLabel.ForeColor = Color.DarkGoldenrod;
                statusLabel.Text = "Chưa có gói pháp lý sẵn sàng: " + ex.Message +
                    " Module vẫn mở; hãy chọn/cài gói dữ liệu để tra và gắn định mức.";
            }
        }

        private void RefreshStateMetrics()
        {
            EstimateV2State state;
            if (!WorkbookEstimateV2StateService.TryLoad(workbook, out state))
            {
                boundValue.Text = "0";
                unboundValue.Text = "0";
                totalValue.Text = "0";
                return;
            }

            EstimateV2WorkItemState[] active = state.WorkItems
                .Where(item => !item.IsOrphaned)
                .ToArray();
            boundValue.Text = active.Count(item => item.HasNormBinding).ToString("N0");
            unboundValue.Text = active.Count(item => !item.HasNormBinding).ToString("N0");
            totalValue.Text = active
                .Where(item => item.HasNormBinding)
                .Select(item => item.NormCode + "|" + item.VariantCode)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count()
                .ToString("N0");
        }

        private void RefreshSearch()
        {
            resultGrid.Rows.Clear();
            currentChoices = new NormChoice[0];
            if (searchIndex == null)
                return;

            try
            {
                IReadOnlyList<NormSearchResult> results = searchIndex.Search(
                    new NormSearchQuery(
                        searchBox.Text,
                        environmentFilter,
                        maximumResults: 100));
                var choices = new List<NormChoice>();
                foreach (NormSearchResult result in results)
                {
                    string[] variants = result.Definition != null && result.Definition.Variants.Count > 0
                        ? result.Definition.Variants.ToArray()
                        : new[] { string.Empty };
                    foreach (string variant in variants)
                    {
                        var choice = new NormChoice(result, variant);
                        choices.Add(choice);
                        int rowIndex = resultGrid.Rows.Add(
                            choice.DisplayCode,
                            result.Title,
                            result.WorkUnit);
                        resultGrid.Rows[rowIndex].Tag = choice;
                    }
                }
                currentChoices = choices.AsReadOnly();
                if (resultGrid.Rows.Count > 0)
                    resultGrid.Rows[0].Selected = true;
                RefreshDetail();
            }
            catch (Exception ex)
            {
                statusLabel.BackColor = Color.MistyRose;
                statusLabel.ForeColor = Color.Firebrick;
                statusLabel.Text = "Không tra được định mức: " + ex.Message;
            }
        }

        private void RefreshDetail()
        {
            NormChoice choice = SelectedChoice();
            foreach (TabPage page in detailTabs.TabPages)
            {
                var grid = page.Controls.OfType<DataGridView>().FirstOrDefault();
                if (grid == null)
                    continue;
                grid.Rows.Clear();
                if (choice?.Result.Definition == null)
                    continue;

                int variantIndex = 0;
                if (choice.Result.Definition.Variants.Count > 0 &&
                    choice.VariantCode.Length > 0)
                {
                    variantIndex = choice.Result.Definition.Variants
                        .Select((value, index) => new { value, index })
                        .Where(item => string.Equals(
                            item.value,
                            choice.VariantCode,
                            StringComparison.OrdinalIgnoreCase))
                        .Select(item => item.index)
                        .DefaultIfEmpty(0)
                        .First();
                }

                NormResourceKind kind = (NormResourceKind)page.Tag;
                foreach (NormResourceRate rate in choice.Result.Definition.Rates.Where(rate => rate.Kind == kind))
                {
                    decimal quantity = rate.Quantities.Count > variantIndex
                        ? rate.Quantities[variantIndex]
                        : 0m;
                    grid.Rows.Add(
                        rate.ResourceCode,
                        rate.ResourceCode,
                        rate.Unit,
                        quantity.ToString("0.####", CultureInfo.CurrentCulture));
                }
                page.Text = PageTitle(kind) + " (" + grid.Rows.Count + ")";
            }
        }

        private void Bind(bool allSelectedRows)
        {
            NormChoice choice = SelectedChoice();
            if (choice == null)
            {
                ShowStatus("Hãy chọn một định mức.", true);
                return;
            }

            try
            {
                IReadOnlyList<SelectedWorkItem> selected =
                    ResolveSelectedWorkItems(allSelectedRows);
                if (selected.Count == 0)
                    throw new InvalidOperationException(
                        "Hãy chọn một dòng công tác đã đăng ký trong sheet Gia DT TC.");

                foreach (SelectedWorkItem workItem in selected)
                {
                    WorkbookEstimateV2StateService.BindNorm(
                        workbook,
                        workItem.WorkItemId,
                        choice.Result.Key,
                        choice.VariantCode,
                        choice.Result.PackageId,
                        choice.Result.PackageVersion);
                }

                WorkbookEstimateV2RegistrationService.ReconcileAll(workbook);
                RefreshStateMetrics();
                ShowStatus(
                    "Đã gắn " + choice.DisplayCode + " cho " +
                    selected.Count + " công tác. Binding được lưu trong Custom XML của workbook.",
                    false);
            }
            catch (Exception ex)
            {
                ShowStatus(ex.Message, true);
            }
        }

        private void UnbindSelected()
        {
            try
            {
                IReadOnlyList<SelectedWorkItem> selected = ResolveSelectedWorkItems(true);
                if (selected.Count == 0)
                    throw new InvalidOperationException("Hãy chọn công tác cần bỏ gắn.");
                int changed = 0;
                foreach (SelectedWorkItem workItem in selected)
                {
                    if (WorkbookEstimateV2StateService.UnbindNorm(workbook, workItem.WorkItemId))
                        changed++;
                }
                ClearVisibleNormForSelected(selected);
                WorkbookEstimateV2RegistrationService.ReconcileAll(workbook);
                RefreshStateMetrics();
                ShowStatus("Đã bỏ gắn định mức cho " + changed + " công tác.", false);
            }
            catch (Exception ex)
            {
                ShowStatus(ex.Message, true);
            }
        }

        private IReadOnlyList<SelectedWorkItem> ResolveSelectedWorkItems(bool allRows)
        {
            Excel.Range selection = null;
            Excel.Worksheet activeSheet = null;
            try
            {
                selection = workbook.Application.Selection as Excel.Range;
                if (selection == null)
                    return new SelectedWorkItem[0];
                activeSheet = selection.Worksheet;
                EstimateV2RegisteredSource source = WorkbookEstimateV2RegistrationService
                    .ListRegistered(workbook)
                    .FirstOrDefault(item =>
                        string.Equals(
                            item.WorksheetCodeName,
                            activeSheet.CodeName,
                            StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(
                            item.WorksheetName,
                            activeSheet.Name,
                            StringComparison.OrdinalIgnoreCase));
                if (source == null)
                    return new SelectedWorkItem[0];

                var result = new List<SelectedWorkItem>();
                var seenRows = new HashSet<int>();
                foreach (Excel.Area area in selection.Areas)
                {
                    try
                    {
                        int firstRow = area.Row;
                        int lastRow = area.Row + area.Rows.Count - 1;
                        for (int row = firstRow; row <= lastRow; row++)
                        {
                            if (row < source.FirstDataRow || row > source.LastDataRow ||
                                !seenRows.Add(row))
                                continue;
                            string id = ReadText(
                                activeSheet,
                                row,
                                source.Columns.TechnicalIdColumn);
                            if (!EstimateV2WorkItemState.IsValidId(id))
                                continue;
                            result.Add(new SelectedWorkItem(id, row, source));
                            if (!allRows)
                                return result.AsReadOnly();
                        }
                    }
                    finally
                    {
                        Release(area);
                    }
                }
                return result.AsReadOnly();
            }
            finally
            {
                Release(activeSheet);
                Release(selection);
            }
        }

        private void ClearVisibleNormForSelected(IEnumerable<SelectedWorkItem> selected)
        {
            foreach (SelectedWorkItem item in selected)
            {
                Excel.Worksheet worksheet = null;
                Excel.Range cell = null;
                try
                {
                    worksheet = ResolveWorksheet(item.Source);
                    cell = worksheet.Cells[item.Row, item.Source.Columns.NormDisplayColumn] as Excel.Range;
                    if (cell != null)
                        cell.ClearContents();
                }
                finally
                {
                    Release(cell);
                    Release(worksheet);
                }
            }
        }

        private Excel.Worksheet ResolveWorksheet(EstimateV2RegisteredSource source)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets.Item[index] as Excel.Worksheet;
                        if (sheet != null && (
                            string.Equals(sheet.CodeName, source.WorksheetCodeName, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(sheet.Name, source.WorksheetName, StringComparison.OrdinalIgnoreCase)))
                        {
                            Excel.Worksheet result = sheet;
                            sheet = null;
                            return result;
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }
                throw new InvalidOperationException("Không tìm thấy sheet công tác đã đăng ký.");
            }
            finally
            {
                Release(sheets);
            }
        }

        private static string ReadText(Excel.Worksheet sheet, int row, int column)
        {
            Excel.Range cell = null;
            try
            {
                cell = sheet.Cells[row, column] as Excel.Range;
                return (Convert.ToString(cell?.Value2, CultureInfo.CurrentCulture) ?? string.Empty).Trim();
            }
            finally
            {
                Release(cell);
            }
        }

        private Control BuildHeader()
        {
            var panel = new Panel { Dock = DockStyle.Top, Height = 76, BackColor = Color.White };
            var back = new Button
            {
                Text = "‹",
                Location = new Point(0, 2),
                Size = new Size(32, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = TextMuted,
                Font = new Font("Segoe UI", 15f),
                Cursor = Cursors.Hand
            };
            back.FlatAppearance.BorderSize = 0;
            back.Click += (s, e) => backAction?.Invoke();
            var title = new Label
            {
                Text = "Gắn định mức",
                Location = new Point(39, 5),
                AutoSize = true,
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor = TextDark
            };
            var subtitle = new Label
            {
                Text = "Chọn và gắn định mức cho dòng công tác đang chọn",
                Location = new Point(40, 38),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.4f),
                ForeColor = TextMuted
            };
            panel.Controls.Add(back);
            panel.Controls.Add(title);
            panel.Controls.Add(subtitle);
            return panel;
        }

        private void AddFilter(string text, NormEnvironment? value, bool selected)
        {
            var radio = new RadioButton
            {
                Text = text,
                Tag = value,
                Appearance = Appearance.Button,
                AutoSize = true,
                Height = 29,
                Checked = selected,
                FlatStyle = FlatStyle.Flat,
                BackColor = selected ? Green : Color.FromArgb(245, 247, 248),
                ForeColor = selected ? Color.White : TextDark,
                Font = new Font("Segoe UI", 7.7f),
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0, 0, 6, 0),
                Padding = new Padding(8, 0, 8, 0),
                Cursor = Cursors.Hand
            };
            radio.FlatAppearance.BorderColor = Border;
            radio.CheckedChanged += (s, e) =>
            {
                if (!radio.Checked)
                    return;
                environmentFilter = (NormEnvironment?)radio.Tag;
                foreach (RadioButton other in filterPanel.Controls.OfType<RadioButton>())
                {
                    other.BackColor = other.Checked ? Green : Color.FromArgb(245, 247, 248);
                    other.ForeColor = other.Checked ? Color.White : TextDark;
                }
                RefreshSearch();
            };
            filterPanel.Controls.Add(radio);
        }

        private TabPage CreateDetailPage(string title, NormResourceKind kind)
        {
            var page = new TabPage(title) { Tag = kind, BackColor = Color.White };
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Segoe UI", 7.7f)
            };
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Mã liệu", Width = 82 });
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Tên tài nguyên",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 100
            });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Đơn vị", Width = 58 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Định mức", Width = 70 });
            page.Controls.Add(grid);
            return page;
        }

        private NormChoice SelectedChoice()
        {
            if (resultGrid.SelectedRows.Count == 0)
                return null;
            return resultGrid.SelectedRows[0].Tag as NormChoice;
        }

        private void ShowStatus(string message, bool error)
        {
            statusLabel.BackColor = error ? Color.MistyRose : GreenSoft;
            statusLabel.ForeColor = error ? Color.Firebrick : GreenDark;
            statusLabel.Text = message;
        }

        private static string PageTitle(NormResourceKind kind)
        {
            switch (kind)
            {
                case NormResourceKind.Material: return "Vật liệu";
                case NormResourceKind.Labor: return "Nhân công";
                case NormResourceKind.Machine: return "Máy thi công";
                default: return kind.ToString();
            }
        }

        private static Label AddMetric(
            TableLayoutPanel parent,
            int column,
            string title,
            EstimateUiIconKind iconKind,
            Color color)
        {
            var card = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 7, 0),
                BackColor = Color.White
            };
            card.Paint += (s, e) =>
            {
                using (var pen = new Pen(Border))
                    e.Graphics.DrawRectangle(pen, 0, 0, Math.Max(0, card.Width - 1), Math.Max(0, card.Height - 1));
            };
            card.Controls.Add(new PictureBox
            {
                Image = EstimateUiIcons.Create(iconKind, 23, color),
                Location = new Point(9, 10),
                Size = new Size(28, 28),
                SizeMode = PictureBoxSizeMode.CenterImage
            });
            var titleLabel = new Label
            {
                Text = title,
                Location = new Point(40, 9),
                Height = 24,
                Font = new Font("Segoe UI", 7.4f),
                ForeColor = TextDark,
                AutoEllipsis = true
            };
            var value = new Label
            {
                Location = new Point(8, 39),
                Height = 38,
                Font = new Font("Segoe UI", 16f, FontStyle.Bold),
                ForeColor = color,
                TextAlign = ContentAlignment.MiddleCenter
            };
            card.Controls.Add(titleLabel);
            card.Controls.Add(value);
            card.Resize += (s, e) =>
            {
                titleLabel.Width = Math.Max(45, card.ClientSize.Width - 44);
                value.Width = Math.Max(50, card.ClientSize.Width - 16);
            };
            parent.Controls.Add(card, column, 0);
            return value;
        }

        private static Button ActionButton(string text, Color background, Color foreground)
        {
            var button = new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 5, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = background,
                ForeColor = foreground,
                Font = new Font("Segoe UI", 7.2f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderColor = background == Color.White ? Border : background;
            button.FlatAppearance.BorderSize = 1;
            return button;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }

        private sealed class NormChoice
        {
            internal NormChoice(NormSearchResult result, string variantCode)
            {
                Result = result;
                VariantCode = (variantCode ?? string.Empty).Trim();
            }

            internal NormSearchResult Result { get; }
            internal string VariantCode { get; }
            internal string DisplayCode =>
                VariantCode.Length == 0
                    ? Result.Key
                    : Result.Key + " / " + VariantCode;
        }

        private sealed class SelectedWorkItem
        {
            internal SelectedWorkItem(
                string workItemId,
                int row,
                EstimateV2RegisteredSource source)
            {
                WorkItemId = workItemId;
                Row = row;
                Source = source;
            }

            internal string WorkItemId { get; }
            internal int Row { get; }
            internal EstimateV2RegisteredSource Source { get; }
        }
    }
}
