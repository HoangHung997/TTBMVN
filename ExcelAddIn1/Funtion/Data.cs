using System;
using System.Data;

using NCalc;


namespace ExcelAddIn1.Funtion
{
    public  class Data
    {
       private  DataTable dt;
        public  Data()
        {
             dt = new DataTable();
            dt.Columns.Add("10", typeof(double));
            dt.Columns.Add("20", typeof(double));
            dt.Columns.Add("50", typeof(double));
            dt.Columns.Add("100", typeof(double));
            dt.Columns.Add("200", typeof(double));

            dt.Columns.Add("500", typeof(double));

            dt.Columns.Add("1000", typeof(double));

            dt.Columns.Add("2000", typeof(double));

            dt.Columns.Add("5000", typeof(double));


            dt.Columns.Add("8000", typeof(double));

            dt.Columns.Add("10000", typeof(double));
            dt.Rows.Add(3.285, 2.853, 2.435, 1.845, 1.546, 1.188, 0.797, 0.694, 0.62, 0.53, 0.478); // Hàng đầu tiên để kiểm tra khoảng
            dt.Rows.Add(3.508, 3.137, 2.559, 2.074, 1.604, 1.301, 0.823, 0.716, 0.64, 0.55, 0.493);
            dt.Rows.Add(3.203, 2.7, 2.356, 1.714, 1.272, 1.003, 0.731, 0.636, 0.55, 0.48, 0.438);
            dt.Rows.Add(2.598, 2.292, 2.075, 1.545, 1.189, 0.95, 0.631, 0.55, 0.49, 0.42, 0.378);
            dt.Rows.Add(2.566, 2.256, 1.984, 1.461, 1.142, 0.912, 0.584, 0.509, 0.452, 0.39, 0.35);
        }
            public  object GetCellValueGS(double value, double rowIndex1, string userFormula)
        {
            // Kiểm tra các tham số hợp lệ
            int rowIndex = Convert.ToInt32(rowIndex1);
            if (dt == null || dt.Rows.Count < 2 || rowIndex < 0 || rowIndex >= dt.Rows.Count)
                throw new ArgumentException("Dữ liệu không hợp lệ hoặc rowIndex ngoài phạm vi.");

            double inputValue = Convert.ToDouble(value);

            // Duyệt qua các cột của hàng đầu tiên
            for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
            {
                var cellValue = Convert.ToDouble(dt.Rows[0][colIndex]);

                // Nếu giá trị khớp chính xác, trả về giá trị tương ứng
                if (cellValue == inputValue)
                {
                    return dt.Rows[rowIndex][colIndex];
                    double valueC = Convert.ToDouble(cellValue);
                }

                // Nếu không phải cột cuối cùng, kiểm tra khoảng giữa hai cột liên tiếp
                if (colIndex < dt.Columns.Count - 1)
                {
                    var nextCellValue = Convert.ToDouble(dt.Rows[0][colIndex + 1]);
                    double valueD = Convert.ToDouble(nextCellValue);

                    // Kiểm tra giá trị nằm trong khoảng
                    if (inputValue > cellValue && inputValue < nextCellValue)
                    {
                        
                        // Tính toán dựa trên công thức người dùng nhập
                        double valueA = Convert.ToDouble(dt.Rows[rowIndex][colIndex]);
                        double valueB = Convert.ToDouble(dt.Rows[rowIndex][colIndex + 1]);

                        // Sử dụng công thức
                        Expression formula = new Expression(userFormula);
                        formula.Parameters["A"] = valueA;
                        formula.Parameters["B"] = valueB;
                        formula.Parameters["C"] = valueA;
                        formula.Parameters["D"] = valueA;
                          formula.Parameters["giatri"] = value;
            return formula.Evaluate();
                    }
                }
            }

            // Nếu không tìm thấy, trả về null hoặc giá trị mặc định
            return null;
        }
    }
}
