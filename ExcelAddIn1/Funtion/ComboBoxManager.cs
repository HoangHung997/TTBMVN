using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq; 

namespace ExcelAddIn1.Funtion
{
    public class ComboBoxManager
    {
        // Dictionary lưu trữ nhiều danh sách dữ liệu
        private Dictionary<string, List<KeyValuePair<string, double>>> danhSachDict;
       
        

        // Khởi tạo lớp với dữ liệu mẫu
        public ComboBoxManager()
        {
            // Tạo và lưu trữ nhiều danh sách vào Dictionary
            // Khởi tạo danhSachTongHop
     
            danhSachDict = new Dictionary<string, List<KeyValuePair<string, double>>>()
            {
                { "cbLoaidutoan", new List<KeyValuePair<string, double >>()
                    {
                        new KeyValuePair<string, double >("Sinh hoạt phí", 1),
                        new KeyValuePair<string, double >("Doanh nghiệp", 2),
                        new KeyValuePair<string, double >("Sinh hoạt phí dự kiến", 3),
                        new KeyValuePair<string, double>("Doanh nghiệp dự kiến", 4)
                    }
                },
                { "cbLoaidiahinh", new List<KeyValuePair<string, double >>()
                    {
                        new KeyValuePair<string, double>("Đồng bằng, trống trải", 2),
                        new KeyValuePair<string, double>("Đô thị, khu dân cư", 2.5),
                        new KeyValuePair<string, double>("Trung du hoặc rừng loại 1", 3),
                        new KeyValuePair<string, double>("Rừng loại 2", 3.5),
                        new KeyValuePair<string, double>("Rừng loại 3", 4),
                        new KeyValuePair<string, double>("Rừng loại 4", 4.5),
                        new KeyValuePair<string, double>("Dưới nước", 3),
                        new KeyValuePair<string, double>("Dưới biển", 5),
                    }
                },
                { "cbKhoiluonghuy", new List<KeyValuePair<string, double >>()
                    {
                        new KeyValuePair<string, double>("Khối lượng dưới 1000 Kg", 5),
                        new KeyValuePair<string, double>("Khối lượng trên 1000 Kg", 3)
                    }
                },
                { "cbLantrai", new List<KeyValuePair<string, double >>()
                    {
                        new KeyValuePair<string, double >("Công trình xây dựng theo tuyến", 1),
                        new KeyValuePair<string, double >("Công trình xây dựng còn lại", 2)
                    }

                },
                { "cbGiamsat", new List<KeyValuePair<string, double >>()
                    {
                         new KeyValuePair<string, double>("Không có giám sát", 0),
                        new KeyValuePair<string, double>("Công trình dân dụng", 1),
                        new KeyValuePair<string, double>("Công trình công nghiệp", 2),
                        new KeyValuePair<string, double>("Công trình giao thông", 3),
                        new KeyValuePair<string, double>("Công trình NN&PTNT", 4),
                        new KeyValuePair<string, double>("Công trình hạ tầng kỹ thuật", 5),


                    }
                }
                // Bạn có thể thêm nhiều danh sách khác vào đây...
            };
           
        }

        // Liên kết dữ liệu vào ComboBox theo tên danh sách
        public void BindComboBox(ComboBox comboBox, string comboBoxName)
        {
            if (danhSachDict.ContainsKey(comboBoxName))
            {
                var danhSach = danhSachDict[comboBoxName];
                comboBox.DisplayMember = "Key";  // Tên hiển thị
                comboBox.ValueMember = "Value";  // Giá trị thực tế
                comboBox.DataSource = danhSach;
            }
            else
            {
                MessageBox.Show("Danh sách không tồn tại");
            }
        }
        // Lấy giá trị khi người dùng chọn item (giá trị double)
        public double GetSelectedValue(ComboBox comboBox)
        {
            if (comboBox.SelectedItem != null)
            {
                return ((KeyValuePair<string, double>)comboBox.SelectedItem).Value;
            }

            return -1;  // Nếu không có giá trị nào được chọn
        }
        // Hàm tra cứu giá trị dựa trên Dictionary tên, Nhóm và Key

        
    }
}
