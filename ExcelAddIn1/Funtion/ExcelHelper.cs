using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public class ExcelHelper
    {
        private Excel.Application excelApp;
        private Excel.Worksheet activeWorksheet;

        public ExcelHelper()
        {
            // Khởi tạo Excel application và worksheet đang hoạt động
            excelApp = Globals.ThisAddIn.Application;
            activeWorksheet = excelApp.ActiveSheet;
        }

        public Excel.Range GetRangeFromAddress(string address)
        {
            try
            {
                // Chuyển địa chỉ ô từ chuỗi thành đối tượng Range
                Excel.Range range = activeWorksheet.get_Range(address, Type.Missing);
                return range;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Đã xảy ra lỗi: " + ex.Message, "Lỗi");
                return null;
            }
        }
    }

}
