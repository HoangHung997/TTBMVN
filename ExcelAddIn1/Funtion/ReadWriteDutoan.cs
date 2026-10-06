using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn1.Funtion
{
    public static class ReadWriteDutoan

    {// Khai báo danh sách các giá trị

        public static void SaveVars(string propertyName, string propertyValue)
        {
            var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var customProperties = workbook.CustomDocumentProperties;

            try
            {
                // Kiểm tra xem property có tồn tại không
                bool exists = false;
                foreach (dynamic prop in customProperties)
                {
                    if (prop.Name == propertyName)
                    {
                        prop.Value = propertyValue; // Cập nhật giá trị nếu tồn tại
                        exists = true;
                        break;
                    }
                }

                // Nếu chưa tồn tại, thêm mới
                if (!exists)
                {
                    customProperties.Add(propertyName, false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, propertyValue);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi lưu giá trị vào CustomDocumentProperties: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        public static string LoadVars(string propertyName)
        {
            var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var customProperties = workbook.CustomDocumentProperties;

            try
            {
                // Duyệt qua từng property và so sánh
                foreach (dynamic prop in customProperties)
                {
                    if (prop.Name == propertyName)
                    {
                        return prop.Value.ToString();
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi đọc giá trị từ CustomDocumentProperties: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

    }
}
