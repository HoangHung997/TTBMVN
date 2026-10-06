using ExcelAddIn1.Funtion;
using ExcelAddIn1.Winform;
using ExcelAddIn1.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
            public Boolean CheckIfEditing = false;
        // Giu mot cua so modeless cho moi workbook dang mo.
        private readonly Dictionary<string, Dutoan> openedForms =
            new Dictionary<string, Dutoan>(StringComparer.Ordinal);

        private sealed class ExcelWindowHandle : IWin32Window
        {
            public ExcelWindowHandle(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            EdCottinhtoan.Text = "H";
            edPrefix.Text = "Bằng chữ: ";
            edSuffix.Text = "./.";
            edVND.Text = "E25";

            btnEstimateSettings.Image =
                EstimateUiIcons.Create(
                    EstimateUiIconKind.Settings,
                    32,
                    System.Drawing.Color.FromArgb(
                        0, 103, 55));
            btnEstimateSettings.ShowImage = true;

            // Đăng ký sự kiện Workbook

            //Excel.Application excelApp = Globals.ThisAddIn.Application;
            //excelApp.WorkbookOpen += ExcelApp_WorkbookOpen;
            // Tìm menu bạn đã thêm


            // Thêm các item vào menu





        }

        private void ExcelApp_WorkbookOpen(Excel.Workbook Wb)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            // Đăng ký sự kiện SheetChange và SheetCalculate
            workbook.SheetChange += Workbook_SheetChange;

            workbook.SheetBeforeDoubleClick += Workbook_SheetBeforeDoubleClick;

            //workbook.SheetSelectionChange += Workbook_SheetSelectionChange;
            workbook.SheetActivate += Workbook_SheetActivate;
            LoadSheetNames();
        }

        private void Workbook_SheetActivate(object Sh)
        {
            LoadSheetNames();
        }

        //private void Workbook_SheetSelectionChange(object Sh, Range Target)
        //{
        //    LoadSheetNames();
        //    MessageBox.Show("hung");
        //}

 

        private void Workbook_SheetChange(object Sh, Range Target)
        {
            if (CheckIfEditing)
            {
                // Việc chỉnh sửa đã hoàn tất
                CheckIfEditing = false;
            }
            LoadSheetNames();

        }


        private void Workbook_SheetBeforeDoubleClick(object Sh, Range Target, ref bool Cancel)
        {
            CheckIfEditing = true;
        }



        private void LoadSheetNames()
        {
            {
                // Lấy danh sách tên các sheet
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Workbook workbook = excelApp.ActiveWorkbook;

                if (workbook != null)
                {
                    // Xóa các mục hiện tại nếu có
                    comboBox1.Items.Clear();

                    // Thêm tên các sheet vào ComboBox
                    foreach (Excel.Worksheet sheet in workbook.Sheets)
                    {
                        RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                        item.Label = sheet.Name;
                        comboBox1.Items.Add(item);
                    }
                }
            }
        }
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (CheckIfEditing)
            {
                return;
            }  
            string ColumnSheetTT = EdCottinhtoan.Text;

            TinhToan.DonGia("Vị trí số hiệu định mức", ColumnSheetTT);

         }


        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // Xử lý khi người dùng thay đổi lựa chọn trong ComboBox
            RibbonComboBox comboBox = (RibbonComboBox)sender;
            string selectedValue = comboBox.Text;

            // Kiểm tra nếu giá trị chọn không có trong danh sách
            bool isValid = false;
            foreach (RibbonDropDownItem item in comboBox.Items)
            {
                if (item.Label == selectedValue)
                {
                    isValid = true;
                    break;
                }
            }

            if (!isValid)
            {
                // Nếu giá trị không hợp lệ, đặt lại giá trị chọn
                // Hoặc thông báo cho người dùng
                comboBox.Text = comboBox.Items[0].Label;
                // Thay thế dòng dưới bằng cách thông báo cho người dùng nếu cần
                // MessageBox.Show("Giá trị bạn nhập không hợp lệ.");
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {


            MessageBox.Show(CheckIfEditing.ToString());
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {

            ExcelAddIn1.RandomDaodat.Daodat(false,chkRownew.Checked);
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //TinhToan.Calculate(1, true);
            ExcelAddIn1.RandomDaodat.Daodat(true, chkRownew.Checked);
        }

        private void btPickO_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeWorksheet = excelApp.ActiveSheet;

            // Hiển thị hộp thoại để người dùng nhập địa chỉ ô
            object input = excelApp.InputBox("Chọn một ô bằng cách nhập địa chỉ của nó (ví dụ: A1):", "Chọn 1 ô", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
           
            
            if (input is Excel.Range cellAddress)
            {
                try
                {
                    string cellstring = cellAddress.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                    edVND.Text = cellstring;
                }
                catch (COMException)
                {
                    System.Windows.Forms.MessageBox.Show("Địa chỉ ô không hợp lệ. Vui lòng thử lại.", "Lỗi");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Không có ô nào được chọn", "Thông báo");
            }
        }

        private void btChuyenDoi_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelper helper = new ExcelHelper();

            // Đổi sang icon động khi bắt đầu xử lý
            btChuyenDoi.Image = Properties.Resources.icons8_sync;

            Excel.Application excelApp = Globals.ThisAddIn.Application;
            string cellstring = edVND.Text;
            string tPrefix = edPrefix.Text;
            string tSuffix = edSuffix.Text;

            try
            {
                if (string.IsNullOrEmpty(cellstring)) return;

                Excel.Range cellRange = helper.GetRangeFromAddress(cellstring);
                var cellValue = cellRange?.Value;

                string VNDChu = "0";
                if (decimal.TryParse(cellValue?.ToString() ?? "0", out decimal number))
                    VNDChu = Tools.Vnd(number, tPrefix, tSuffix);
                else
                    VNDChu = Tools.Vnd(0, tPrefix, tSuffix);

                object input = excelApp.InputBox(
                    "Chọn một ô bằng cách nhập địa chỉ của nó (ví dụ: A1):",
                    "Chọn 1 ô",
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                if (input == null || input is bool b && b == false)
                    return;

                if (input is Excel.Range cellAddress)
                    cellAddress.Value = VNDChu;
                else
                    System.Windows.Forms.MessageBox.Show("Không có ô nào được chọn", "Thông báo");
            }
            catch (COMException)
            {
                System.Windows.Forms.MessageBox.Show("Địa chỉ ô không hợp lệ. Vui lòng thử lại.", "Lỗi");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Lỗi: " + ex.Message, "Thông báo");
            }
            finally
            {
                // 🔁 Dù có lỗi, cancel hay thành công — luôn đổi lại icon tĩnh
                btChuyenDoi.Image = Properties.Resources._9104252_reload_refresh_sync_repeat_icon;
            }

        }

        private void Cong_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
          
            string text = editBox1.Text;
            Excelcontrunhanchia.AddValueToSelectedCell(excelApp, text);
        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var selection = app.Selection as Microsoft.Office.Interop.Excel.Range;

            if (selection != null)
            {
                // Danh sách các ô sẽ nằm trong group
                Microsoft.Office.Interop.Excel.Range groupRange = null;

                foreach (Microsoft.Office.Interop.Excel.Range cell in selection.Cells)
                {
                    if (cell.Value2 != null && cell.Value2.ToString() == "1")
                    {

                        groupRange = app.Union(cell, cell);
                        //if (cell.OutlineLevel > 1)
                        //{
                        //    groupRange.Rows.Ungroup(); // Group các ô nằm trong phạm v
                        //}
                        groupRange.Rows.Group(); // Group các ô nằm trong phạm v
                    }
                }


            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Vui lòng chọn các ô trước khi thực hiện.", "Thông báo");
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var selection = app.Selection as Microsoft.Office.Interop.Excel.Range;

            if (selection != null)
            {
                // Danh sách các ô sẽ nằm trong group
                Microsoft.Office.Interop.Excel.Range groupRange = null;

                foreach (Microsoft.Office.Interop.Excel.Range cell in selection.Cells)
                {
                    if (cell.Value2 != null && cell.Value2.ToString() == "1")
                    {

                        groupRange = app.Union(cell, cell);
                        //if (cell.OutlineLevel > 1)
                        //{
                        //    groupRange.Rows.Ungroup(); // Group các ô nằm trong phạm v
                        //}
                        groupRange.Columns.Group(); // Group các ô nằm trong phạm v
                    }
                }


            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Vui lòng chọn các ô trước khi thực hiện.", "Thông báo");
            }
        }



        private void btnDutoan_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (currentWorkbook == null)
            {
                MessageBox.Show(
                    "Không có file Excel nào đang mở!",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // V2 UI: sử dụng CustomTaskPane dock bên phải theo bộ ảnh UI đã chốt.
                // Không mở form modal/project setup trước, không chiếm vùng làm việc Excel.
                EstimateTaskPaneManager.Show(currentWorkbook);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open estimate task pane");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được Trợ lý Dự toán",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void btnEstimateSettings_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            Excel.Workbook currentWorkbook =
                Globals.ThisAddIn.Application.ActiveWorkbook;
            if (currentWorkbook == null)
            {
                MessageBox.Show(
                    "Không có file Excel nào đang mở!",
                    "Thông báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                EstimateTaskPaneManager.ShowSettings(
                    currentWorkbook);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(
                    ex,
                    "Open Estimate V2 settings");
                MessageBox.Show(
                    ex.Message,
                    "Không mở được Thiết lập chung",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private void btnEstimatePackages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook != null) EstimateTaskPaneManager.ShowPackages(workbook);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open Estimate V2 packages");
                MessageBox.Show(ex.Message, "Gói pháp lý & Dữ liệu");
            }
        }

        private void btnEstimateReports_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook != null) EstimateTaskPaneManager.ShowReports(workbook);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Open Estimate V2 reports");
                MessageBox.Show(ex.Message, "Báo cáo & Xuất in");
            }
        }

        private static string GetWorkbookWindowKey(Excel.Workbook workbook)
        {
            IntPtr unknown = IntPtr.Zero;
            try
            {
                unknown = Marshal.GetIUnknownForObject(workbook);
                return unknown.ToInt64().ToString("X16", System.Globalization.CultureInfo.InvariantCulture);
            }
            finally
            {
                if (unknown != IntPtr.Zero)
                    Marshal.Release(unknown);
            }
        }

        private void btnRandom3m_Click(object sender, RibbonControlEventArgs e)
        {
            TinhToan.Random(1, false);

        }

        private void btnRandom5m_Click(object sender, RibbonControlEventArgs e)
        {
            TinhToan.Random(1, true);
        }

        private void btnThrd_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            IntPtr excelHandle = new IntPtr(Globals.ThisAddIn.Application.Hwnd);
            NativeWindow excelWindow = new NativeWindow();
            excelWindow.AssignHandle(excelHandle);

            FrmDaodat form = new FrmDaodat(chkRownew.Checked);
            form.Show(excelWindow);
        }

        private void group2_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            IntPtr excelHandle = new IntPtr(Globals.ThisAddIn.Application.Hwnd);
            NativeWindow excelWindow = new NativeWindow();
            excelWindow.AssignHandle(excelHandle);

            FrmDaodat form = new FrmDaodat(chkRownew.Checked);
            form.Show(excelWindow);
        }
    }


}
