using ExcelAddIn1.Funtion;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReadWriteDutoan = ExcelAddIn1.Funtion.ReadWriteDutoan;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Globalization;


namespace ExcelAddIn1.Winform
{
    public partial class FormThongtinchung : Form
    {
        private ComboBoxManager comboBoxManager;
        public FormThongtinchung()
        {
            InitializeComponent();
            this.AutoScaleMode = AutoScaleMode.Dpi;
            CaiDatDuToanForm();



            tbGiatridutoan.Enabled = false;
            tbGiatrỉPBM.Enabled = false;

            // Khởi tạo đối tượng ComboBoxManager

            comboBoxManager = new ComboBoxManager();
            // Liên kết dữ liệu vào ComboBox với tên danh sách
            comboBoxManager.BindComboBox(cbLoaidutoan, "cbLoaidutoan");
            comboBoxManager.BindComboBox(cbLoaidiahinh, "cbLoaidiahinh");
            comboBoxManager.BindComboBox(cbKhoiluonghuy, "cbKhoiluonghuy");
            comboBoxManager.BindComboBox(cbLantrai, "cbLantrai");
            comboBoxManager.BindComboBox(cbGiamsat, "cbGiamsat");
            loadSetting();
            //UpdateTextbox();

        }
        private void loadSetting()
        {
            LoadComboBoxSettings(cbLoaidutoan, cbLoaidiahinh, cbKhoiluonghuy, cbLantrai, cbGiamsat);

        }
        private void CaiDatDuToanForm()
        {
            // Gán tag name cho từng nút (đặt theo ý bạn)
            SetupButton(button1, "LoaiDuToan");
            SetupButton(button2, "LoaiDiaHinh");
            SetupButton(button3, "KhoiLuongHuyNo");
            SetupButton(button4, "LoaiCongTrinhLanTrai");
            SetupButton(button5, "LoaiCongTrinhGS");
            SetupButton(button6, "GiaTriRPBM");
        }
        private void SetupButton(System.Windows.Forms.Button btn, string tagName)
        {
            btn.Tag = tagName;
            btn.Click -= TagButton_Click;
            btn.Click += TagButton_Click;
            RefreshButtonCaption(btn);
        }
        private void RefreshButtonCaption(System.Windows.Forms.Button btn)
        {
            string tagName = (string)btn.Tag;
            var rng = CellTagWithShape.GetTaggedRange(tagName);
            btn.Text = rng == null ? "null" : rng.Address[RowAbsolute: false, ColumnAbsolute: false];
        }
        private void TagButton_Click(object sender, EventArgs e)
        {
            var btn = (System.Windows.Forms.Button)sender;
            string tagName = (string)btn.Tag;

            var app = Globals.ThisAddIn.Application;
            // Yêu cầu người dùng chọn 1 ô
            Excel.Range picked = null;
            try
            {
                picked = app.InputBox($"Chọn 1 ô cho '{tagName}'", "Gắn/đổi cell", Type: 8) as Excel.Range;
            }
            catch { /* user Esc */ }

            if (picked == null) return;

            // Nếu đã có shape → di chuyển; nếu chưa → gắn mới
            var current = CellTagWithShape.GetTaggedRange(tagName);
            if (current == null)
                CellTagWithShape.AddTag(tagName, picked);
            else
                CellTagWithShape.MoveTag(tagName, picked);

            RefreshButtonCaption(btn);
        }
        private void LoadComboBoxSettings(params ComboBox[] comboBoxes)
        {
            foreach (var comboBox in comboBoxes)
            {
                // Lấy tên ComboBox làm key
                string savedValue = ReadWriteDutoan.LoadVars(comboBox.Name);

                if (string.IsNullOrEmpty(savedValue) || !int.TryParse(savedValue, out int selectedIndex))
                {

                    // Nếu không có giá trị lưu hoặc giá trị không hợp lệ
                    comboBox.SelectedIndex = comboBox.Items.Count > 0 ? 0 : -1; // Chọn mục đầu tiên nếu danh sách có
                }
                else
                {
                    // Gán giá trị hợp lệ
                    comboBox.SelectedIndex = selectedIndex;
                }
            }
        }
        private void SaveComboBoxSettings(params ComboBox[] comboBoxes)
        {
            foreach (var comboBox in comboBoxes)
            {
                if (comboBox.SelectedIndex >= 0) // Đảm bảo có mục được chọn
                {
                    ReadWriteDutoan.SaveVars(comboBox.Name, comboBox.SelectedIndex.ToString());
                }
            }
        }
        //private void UpdateTextbox()
        //{
        //    Double selectedValue = comboBoxManager.GetSelectedValue(cbLoaidiahinh);
        //    var excelWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Excel.Worksheet targetSheet = (Excel.Worksheet)excelWorkbook.Sheets["THKP-TC"];
        //    Excel.Range columnRange = targetSheet.Range["H:H"];
        //    int Row = ExcelFunctions.FindRow("Loaidiahinh", columnRange);
        //    Range Cell = targetSheet.Range["E" + Row];
        //    Cell.Value2 = selectedValue;


        //    int RowRPBM = ExcelFunctions.FindRow("GiatriRPBM", columnRange);
        //    string formattedNumber = FormatNumberWithDot(targetSheet.Range["F" + RowRPBM].Value2, "#,##0");
        //    tbGiatrỉPBM.Text = formattedNumber;
        //    Row = ExcelFunctions.FindRow("Thamdinh", columnRange);
        //     Cell = targetSheet.Range["E" + Row];
        //    object result = Thamdinh(targetSheet.Range["F" + RowRPBM].Value2);
        //    Cell.Value2 = result;






        //    int DTRow = ExcelFunctions.FindRow("Giatridutoan", columnRange);
        //     formattedNumber = FormatNumberWithDot(targetSheet.Range["F" + DTRow].Value2, "#,##0");
        //    tbGiatridutoan.Text = " Giá trị dự toán: " + formattedNumber;
        //}
        //private void UpdateGiamSat ()
        //{
        //    var excelWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Excel.Worksheet targetSheet = (Excel.Worksheet)excelWorkbook.Sheets["THKP-TC"];
        //    Excel.Range columnRange = targetSheet.Range["H:H"];

        //    int RowGS = ExcelFunctions.FindRow("Giamsat", columnRange);
        //    if (cbGiamsat.SelectedIndex == 0) 
        //     {
        //        targetSheet.Range["E" + RowGS].Value2 = 0;

        //    } 
        //    else
        //    {
        //        int RowRPBM = ExcelFunctions.FindRow("GiatriRPBM", columnRange);
        //        Double RPBM = targetSheet.Range["F" + RowRPBM].Value2;



        //        var result1 = dataDutoan.GetCellValueGS(RPBM, cbGiamsat.SelectedIndex - 1);

        //        targetSheet.Range["E" + RowGS].Value2 = result1;
        //    }    





        //}
        //private void UpdateLantrai()
        //{
        //    var excelWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Excel.Worksheet targetSheet = (Excel.Worksheet)excelWorkbook.Sheets["THKP-TC"];
        //    Excel.Range columnRange = targetSheet.Range["H:H"];
        //    int RowLT = ExcelFunctions.FindRow("Lantrai", columnRange);
        //    if (cbLantrai.SelectedIndex >= 0)
        //    {
        //        int RowRPBM = ExcelFunctions.FindRow("GiatriRPBM", columnRange);
        //        Double RPBM = targetSheet.Range["F" + RowRPBM].Value2;



        //        var result1 = dataDutoan.GetCellValueLantrai(RPBM, cbLantrai.SelectedIndex);

        //        targetSheet.Range["E" + RowLT].Value2 = result1;

        //    }






        //}
        //private object Thamdinh(double GiatriRPBM)
        //{
        //    double giatri = GiatriRPBM / Math.Pow(10, 9);
        //    if (giatri <= 1)
        //    {
        //        if ((GiatriRPBM * 0.005) < 2000000)
        //        {
        //            return "Mức tối thiểu";
        //        }
        //        else if ((GiatriRPBM * 0.005) > 60000000)
        //        {
        //            return "Mức tối đa";
        //        }
        //        else
        //        {
        //            return 0.5; // Sửa từ 0,5 thành 0.5
        //        }
        //    }
        //    else if (giatri > 1 && giatri <= 5)
        //    {
        //        if ((GiatriRPBM * 0.003) < 2000000)
        //        {
        //            return "Mức tối thiểu";
        //        }
        //        else if ((GiatriRPBM * 0.003) > 60000000)
        //        {
        //            return "Mức tối đa";
        //        }
        //        else
        //        {
        //            return 0.3; // Sửa từ 0,3 thành 0.3
        //        }
        //    }
        //    else
        //    {
        //        if ((GiatriRPBM * 0.002) < 2000000)
        //        {
        //            return "Mức tối thiểu";
        //        }
        //        else if ((GiatriRPBM * 0.002) > 60000000)
        //        {
        //            return "Mức tối đa";
        //        }
        //        else
        //        {
        //            return 0.2; // Sửa từ 0,2 thành 0.2
        //        }
        //    }
        //}

        //private void cbLoaidutoan_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (cbLoaidutoan.SelectedIndex > 0) { SaveComboBoxSettings(cbLoaidutoan); }
        //    Double selectedValue = comboBoxManager.GetSelectedValue(cbLoaidutoan);
        //    var excelWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Excel.Worksheet targetSheet = (Excel.Worksheet)excelWorkbook.Sheets["Gia DT TC_DN"];
        //    Excel.Range columnRangePAchon = targetSheet.Range["P:P"];
        //    int RowPAchon = ExcelFunctions.FindRow("PAchon", columnRangePAchon) + 1;
        //    Range CellPAchon = targetSheet.Range["P" + RowPAchon ];

        //    Excel.Worksheet targetSheetLuong = (Excel.Worksheet)excelWorkbook.Sheets["VL-NC-M"];
        //    Excel.Range columnRangeLuong = targetSheetLuong.Range["G:G"];
        //    int RowLuong = ExcelFunctions.FindRow("Luong", columnRangeLuong);
        //    Range CellLuong = targetSheetLuong.Range["D" + RowLuong];

        //    switch (selectedValue)
        //    {
        //        case 1:
        //            CellPAchon.Value2 = 1;
        //            CellLuong.Value2 = 450000;
        //            break;
        //        case 2:
        //            CellPAchon.Value2 = 1;
        //            CellLuong.Value2 = 2340000;
        //            break;
        //        case 3:
        //            CellPAchon.Value2 = 2;
        //            CellLuong.Value2 = 450000;
        //            break;
        //        case 4:
        //            CellPAchon.Value2 = 2;
        //            CellLuong.Value2 = 2340000;
        //            break;
        //    }

        //    // tbGiatridutoan.Text = " cbLoaidutoan: " + selectedValue.ToString();

        //    Excel.Worksheet targetSheetTHKP = (Excel.Worksheet)excelWorkbook.Sheets["THKP-TC"];
        //    Excel.Range columnRangeTHKP = targetSheetTHKP.Range["H:H"];
        //    int RowLDH = ExcelFunctions.FindRow("Loaidiahinh", columnRangeTHKP);
        //    Range Cell = targetSheet.Range["E" + RowLDH];
        //    Cell.Value2 = selectedValue;
        //    UpdateTextbox();
        //}

        //private void cbLoaidiahinh_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (cbLoaidiahinh.SelectedIndex > 0) { SaveComboBoxSettings(cbLoaidiahinh); }
        //    Double selectedValue = comboBoxManager.GetSelectedValue(cbLoaidiahinh);
        //    var excelWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        //    Excel.Worksheet targetSheet = (Excel.Worksheet)excelWorkbook.Sheets["THKP-TC"];
        //    Excel.Range columnRange = targetSheet.Range["H:H"];
        //    int Row = ExcelFunctions.FindRow("Loaidiahinh", columnRange);
        //    Range Cell = targetSheet.Range["E" + Row];
        //    Cell.Value2 = selectedValue;

        //    UpdateTextbox();
        //    // tbGiatridutoan.Text = " cbLoaidiahinh: " + selectedValue.ToString();
        //}

        private void cbKhoiluonghuy_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbKhoiluonghuy.SelectedIndex > 0) { SaveComboBoxSettings(cbKhoiluonghuy); }

            Double selectedValue = comboBoxManager.GetSelectedValue(cbKhoiluonghuy);
            //var excelWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            //Excel.Worksheet targetSheet = (Excel.Worksheet)excelWorkbook.Sheets["THKP-TC"];
            //Excel.Range columnRange = targetSheet.Range["H:H"];
            //int Row = ExcelFunctions.FindRow("Khoiluonghuy", columnRange);
            //Range Cell = targetSheet.Range["E" + Row];
            Range Cell = CellTagWithShape.GetTaggedRange("KhoiLuongHuyNo");
            if (Cell != null)
                Cell.Value2 = selectedValue;


            //UpdateTextbox();
        }

        //private void cbLantrai_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (cbLantrai.SelectedIndex > 0) { SaveComboBoxSettings(cbLantrai); }
        //    UpdateLantrai();
        //    UpdateTextbox();

        //}

        //private void cbGiamsat_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (cbGiamsat.SelectedIndex > 0) { SaveComboBoxSettings(cbGiamsat); }
        //    UpdateGiamSat();
        //    UpdateTextbox();

        //} 
        private string FormatNumberWithDot(object number, string formatNum)
        {
            // Kiểm tra nếu number có thể chuyển đổi thành kiểu số hợp lệ (decimal hoặc double)
            if (number is decimal || number is double)
            {
                // Chuyển đổi về kiểu decimal hoặc double
                decimal num = Convert.ToDecimal(number);

                string formattedNumber = num.ToString(formatNum, CultureInfo.CurrentCulture);

                return formattedNumber;
            }
            else
            {
                // Nếu không phải kiểu số hợp lệ, trả về chuỗi trống hoặc thông báo lỗi
                return "Invalid number";
            }
        }
    }
}
