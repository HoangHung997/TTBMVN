using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public class TinhToan
    {
       

        public static void Calculate(int skipValue, bool is5m)
        {
            var excelApp = Globals.ThisAddIn.Application;
            var worksheet = excelApp.ActiveSheet as Excel.Worksheet;
            var selectedRange = excelApp.Selection as Excel.Range;
            var calculator = new ExcelCalculator(worksheet);
            double khoiLuong;
            int newSkipValue = 0;

            // Kiểm tra nếu các ô được chọn không nằm cùng một cột
            int initialColumn = selectedRange.Cells[1, 1].Column;
            foreach (Excel.Range cell in selectedRange)
            {
                if (cell.Column != initialColumn)
                {
                    MessageBox.Show("Chỉ được chọn các ô trong 1 cột, vui lòng chọn lại", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; // Thoát khỏi phương thức nếu các ô không nằm cùng một cột
                }
            }

            // Xác định ô cuối cùng có dữ liệu trong cột của phạm vi đã chọn
            int nCol = selectedRange.Column;
            Excel.Range lastCellInCol = worksheet.Cells[worksheet.Rows.Count, nCol].End(Excel.XlDirection.xlUp);
            int lastRowInCol = lastCellInCol.Row;

            foreach (Excel.Range activeCell in selectedRange)
            {
                int nRow = activeCell.Row;

                // Nếu dòng hiện tại vượt quá dòng cuối cùng có dữ liệu thì dừng lại
                if (nRow > lastRowInCol)
                {
                    break;
                }

                // Chuyển đổi giá trị ô thành double để đảm bảo tính tương thích
                double cellValue = 0;
                if (activeCell.Value2 != null && double.TryParse(activeCell.Value2.ToString(), out cellValue))
                {
                    if (cellValue == 0)
                    {
                        continue; // Bỏ qua ô có giá trị bằng 0
                    }

                    khoiLuong = Math.Max(0, cellValue);

                    if (newSkipValue > 0)
                    {
                        newSkipValue--;
                        continue;
                    }

                    // Kiểm tra giá trị của ô bên trái
                    Excel.Range leftCell = worksheet.Cells[nRow, nCol - 1];
                    if (leftCell != null && leftCell.Value2 != null && int.TryParse(leftCell.Value2.ToString(), out skipValue))
                    {
                        newSkipValue = skipValue;
                    }

                    // Gọi phương thức xử lý của calculator
                    if (is5m)
                    {
                        calculator.ProcessCell(nRow, nCol, khoiLuong, skipValue, true);
                    }
                    else
                    {
                        calculator.ProcessCell(nRow, nCol, khoiLuong, skipValue, false);
                    }
                }
                else
                {
                    // Nếu giá trị không hợp lệ, bỏ qua ô này
                    continue;
                }
            }
        }
        public static double[] RandomNSoTong(double tong, int soLuong, double min, double max)
        {
            if (soLuong <= 0)
                throw new ArgumentException("Số lượng phải > 0");

            // Làm tròn 2 chữ số → số lượng giá trị khả dụng
            int totalPossibleValues = (int)Math.Floor((max - min) * 100) + 1;
            if (soLuong > totalPossibleValues)
                throw new ArgumentException("Không đủ giá trị khác nhau trong khoảng để tạo ra " + soLuong + " số.");

            // Danh sách tất cả các giá trị khả dụng
            var possibleValues = Enumerable.Range(0, totalPossibleValues)
                                           .Select(i => Math.Round(min + i * 0.01, 2))
                                           .ToList();

            var rnd = new Random();
            for (int attempt = 0; attempt < 10000; attempt++)
            {
                // Bước 1: Random N số khác nhau
                var selected = possibleValues.OrderBy(_ => rnd.Next()).Take(soLuong).ToArray();

                // Bước 2: Scale tổng nếu cần
                double sumSelected = selected.Sum();
                double scale = tong / sumSelected;

                var scaled = selected.Select(x => Math.Round(x * scale, 2)).ToArray();

                // Bước 3: Kiểm tra lại
                if (scaled.All(x => x >= min && x <= max) &&
                    scaled.Distinct().Count() == soLuong &&
                    Math.Abs(scaled.Sum() - tong) <= 0.01)
                {
                    return scaled;
                }
            }

            throw new Exception("Không thể tạo danh sách thỏa mãn sau 10000 lần thử.");
        }
        public static void Random(int skipValue, bool is5m)
        {
            Random random = new Random();
            var excelApp = Globals.ThisAddIn.Application;
            var selectedRange = excelApp.Selection as Excel.Range;
            if (selectedRange == null || selectedRange.Columns.Count != 1)
            {
                MessageBox.Show("Chỉ được chọn các ô trong 1 cột, vui lòng chọn lại", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (new ExcelWriteContext(excelApp))
            using (var transaction = new ExcelBatchWriteTransaction())
            {
                Excel.Areas areas = null;
                try
                {
                    areas = selectedRange.Areas;
                    for (int areaIndex = 1; areaIndex <= areas.Count; areaIndex++)
                    {
                        Excel.Range area = null;
                        Excel.Range output = null;
                        try
                        {
                            area = areas.Item[areaIndex];
                            output = area.Offset[0, 1];
                            int rows = area.Rows.Count;
                            object[,] inputValues = ToMatrix(area.Value2, rows, 1);
                            object[,] formulas = ToMatrix(output.Formula, rows, 1);
                            for (int row = 0; row < rows; row++)
                            {
                                object raw = inputValues[row, 0];
                                if (raw == null || !double.TryParse(raw.ToString(), out double cellValue) || cellValue == 0d)
                                    continue;
                                double factor = is5m
                                    ? 14.67 + random.Next(-20, 51) / 100.0
                                    : 2.86 + random.Next(-20, 21) / 100.0;
                                formulas[row, 0] = factor * cellValue;
                            }
                            transaction.WriteFormula(output, formulas);
                        }
                        finally
                        {
                            ReleaseCom(output);
                            ReleaseCom(area);
                        }
                    }
                }
                finally
                {
                    ReleaseCom(areas);
                }
                transaction.Commit();
            }
        }

        internal static object[,] ToMatrix(object source, int rows, int columns)
        {
            var output = new object[rows, columns];
            Array values = source as Array;
            if (values == null)
            {
                if (rows > 0 && columns > 0)
                    output[0, 0] = source;
                return output;
            }
            int rowLower = values.GetLowerBound(0);
            int columnLower = values.GetLowerBound(1);
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                    output[row, column] = values.GetValue(rowLower + row, columnLower + column);
            }
            return output;
        }

        internal static void ReleaseCom(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
        public static void DonGia(string defaultValue, string ColumnSheetTT)
        {
            ExcelWriteContext writeContext = null;
            try
            {
                SplashManager.ShowForm("Vui lòng chờ!", "Đang xử lý dữ liệu...");
                var excelApp = Globals.ThisAddIn.Application;
                var worksheet = excelApp.ActiveSheet as Excel.Worksheet;
                var selectedRange = excelApp.Selection as Excel.Range;

                writeContext = new ExcelWriteContext(excelApp);

                Excel.Range lastCell = worksheet.Cells[worksheet.Rows.Count, ColumnSheetTT].End(Excel.XlDirection.xlUp);
                int lastRow = lastCell.Row;
                int lastCol = lastCell.Column;

                var excelDongia = new ExcelDongia(worksheet);

                excelDongia.ProcessCellDonGia(lastRow, lastCol, defaultValue, ColumnSheetTT);
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "TinhToan.DonGia", new System.Collections.Generic.Dictionary<string, string>
                {
                    { "ColumnSheetTT", ColumnSheetTT },
                    { "DefaultValue", defaultValue }
                });
                MessageBox.Show("Đã xảy ra lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                writeContext?.Dispose();
                SplashManager.CloseForm();
            }
        }
        public static void RandomTH(Range ranDoi,Range Rano)
        {


        }

    }

    public class ExcelCalculator
    {
        private readonly Excel.Worksheet _worksheet;
        private readonly Random _random = new Random();

        public ExcelCalculator(Excel.Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public void ProcessCell(int nRow, int nCol, double khoiLuong, int skipValue, bool is5m)
        {
            double K4, L4, M4, N4, O4;
            double expression;
            string dau;
            var defaultK4 = is5m ? 3 : 2;
            var defaultL4 = is5m ? 2.5 : 1.5;
            var defaultM4 = is5m ? 1.5 : 1.3;
            var defaultN4 = is5m ? 0.8 : 0.8;

            string numberFormat = _worksheet.Cells[nRow, nCol].numberFormat;
            if (numberFormat.Contains(".") && !numberFormat.Contains(","))
            {
                dau = ".";
            }
            {
                dau = ",";
            }
            _worksheet.Cells[nRow, nCol].NumberFormat = "0"+dau+"00";
            _worksheet.Cells[nRow, nCol - 1].NumberFormat = "0";
            if (skipValue == 1)
            {
                K4 = khoiLuong > 0 ? defaultK4 + _random.Next(-20, 21) / 100.0 : 0;
                L4 = khoiLuong > 0 ? defaultL4 + _random.Next(-10, 11) / 100.0 : 0;
                M4 = khoiLuong > 0 ? defaultM4 + _random.Next(-20, 21) / 100.0 : 0;
                N4 = khoiLuong > 0 ? defaultN4 + _random.Next(-20, 11) / 100.0 : 0;
                O4 = 0;

                expression = khoiLuong * 3 / (K4 * L4 + M4 * N4 + Math.Sqrt(K4 * L4 * M4 * N4));
                O4 = (expression > (is5m ? 3 : 0.3) && expression < (is5m ? 5 : 3)) ? Math.Round(expression, 4) : double.NaN;

                WriteResults(nRow, nCol, K4, L4, M4, N4, O4, is5m, 1 , dau);
                if (_worksheet.Cells[nRow + skipValue, nCol + 5].Value2.ToString() != "Not OK")
                {
                    _worksheet.Cells[nRow, nCol].Value2 = "=" + _worksheet.Cells[nRow + 1, nCol].Address;
                    _worksheet.Cells[nRow, nCol-1].Value2 = "=" + _worksheet.Cells[nRow + 1, nCol-1].Address;
                    if (!_worksheet.Cells[nRow, nCol].Font.Bold)
                    {
                        // Set the text to bold if it is not already bold
                        _worksheet.Cells[nRow, nCol].Font.Bold = true;
                    }
                    if (!_worksheet.Cells[nRow, nCol-1].Font.Bold)
                    {
                        // Set the text to bold if it is not already bold
                        _worksheet.Cells[nRow, nCol-1].Font.Bold = true;
                    }

                }
                else
                {
                    _worksheet.Cells[nRow, nCol].Value2 = khoiLuong;
                }
            }
            else
            {
                ProcessMultipleRows(nRow, nCol, khoiLuong, skipValue, is5m, defaultK4, defaultL4, defaultM4, defaultN4, dau);
            }
        }

        private void ProcessMultipleRows(int nRow, int nCol, double khoiLuong, int skipValue, bool is5m, double defaultK4, double defaultL4, double defaultM4, double defaultN4, string dau)
        {
            double KLmoi = 0;

            for (int i = 1; i < skipValue; i++)
            {
                double K4 = khoiLuong > 0 ? defaultK4 + _random.Next(-20, 20) / 100.0 : 0;
                double L4 = khoiLuong > 0 ? defaultL4 + _random.Next(-10, 10) / 100.0 : 0;
                double M4 = khoiLuong > 0 ? defaultM4 + _random.Next(-20, 20) / 100.0 : 0;
                double N4 = khoiLuong > 0 ? defaultN4 + _random.Next(-20, 10) / 100.0 : 0;
                double J4 = khoiLuong > 0 ? (is5m ? 4 : 1.5) + _random.Next(-50, 50) / 100.0 : 0;

                WriteResults(nRow, nCol, K4, L4, M4, N4, J4, is5m, i, dau);
                KLmoi += _worksheet.Cells[nRow + i, nCol].Value2;
            }
            khoiLuong -= KLmoi;
            
            double K4Final = khoiLuong > 0 ? defaultK4 + _random.Next(-20, 20) / 100.0 : 0;
            double L4Final = khoiLuong > 0 ? defaultL4 + _random.Next(-10, 10) / 100.0 : 0;
            double M4Final = khoiLuong > 0 ? defaultM4 + _random.Next(-20, 20) / 100.0 : 0;
            double N4Final = khoiLuong > 0 ? defaultN4 + _random.Next(-20, 10) / 100.0 : 0;
            double expression = khoiLuong * 3 / (K4Final * L4Final + M4Final * N4Final + Math.Sqrt(K4Final * L4Final * M4Final * N4Final));
            double O4Final = (expression > (is5m ? 3 : 0.3) && expression < (is5m ? 5 : 3)) ? Math.Round(expression, 4) : double.NaN;


            WriteResults(nRow, nCol, K4Final, L4Final, M4Final, N4Final, O4Final, is5m, skipValue, dau);

            if (_worksheet.Cells[nRow + skipValue, nCol + 5].Value2.ToString() != "Not OK")
            {
                _worksheet.Cells[nRow, nCol].Value2 = "=SUM(" + _worksheet.Cells[nRow + 1, nCol].Address + ":" + _worksheet.Cells[nRow + skipValue, nCol].Address + ")";
                _worksheet.Cells[nRow, nCol-1].Value2 = "=SUM(" + _worksheet.Cells[nRow + 1, nCol-1].Address + ":" + _worksheet.Cells[nRow + skipValue, nCol-1].Address + ")";
               
                if (!_worksheet.Cells[nRow, nCol].Font.Bold)
                {
                    // Set the text to bold if it is not already bold
                    _worksheet.Cells[nRow, nCol].Font.Bold = true;  

                }   
                if (!_worksheet.Cells[nRow, nCol-1].Font.Bold)
                {
                    // Set the text to bold if it is not already bold
                    _worksheet.Cells[nRow, nCol-1].Font.Bold = true;  

                }

            }
            else
            {
                _worksheet.Cells[nRow, nCol].Value2 = khoiLuong + KLmoi;
            }
        }

        private void WriteResults(int nRow, int nCol, double K4, double L4, double M4, double N4, double O4, bool is5m, int offset, string dau)
        {
            _worksheet.Cells[nRow + offset, nCol - 1].Value2 = 1;
            _worksheet.Cells[nRow + offset, nCol + 1].Value2 = K4;
            _worksheet.Cells[nRow + offset, nCol + 2].Value2 = L4;
            _worksheet.Cells[nRow + offset, nCol + 3].Value2 = M4;
            _worksheet.Cells[nRow + offset, nCol + 4].Value2 = N4;
            _worksheet.Cells[nRow + offset, nCol + 5].Value2 = double.IsNaN(O4) ? "Not OK" : O4.ToString();
            _worksheet.Cells[nRow + offset, nCol].Value2 = "=" + "ROUND(1/3*" + _worksheet.Cells[nRow + offset, nCol + 5].Address +
                "*((" + _worksheet.Cells[nRow + offset, nCol + 1].Address + "*" + _worksheet.Cells[nRow + offset, nCol + 2].Address +
                ")+(" + _worksheet.Cells[nRow + offset, nCol + 3].Address + "*" + _worksheet.Cells[nRow + offset, nCol + 4].Address +
                ")+SQRT(" + _worksheet.Cells[nRow + offset, nCol + 1].Address + "*" + _worksheet.Cells[nRow + offset, nCol + 2].Address +
                "*" + _worksheet.Cells[nRow + offset, nCol + 3].Address + "*" + _worksheet.Cells[nRow + offset, nCol + 4].Address + ")), 2)";
           
            _worksheet.Cells[nRow + offset, nCol - 1].NumberFormat = "0";
            _worksheet.Cells[nRow + offset, nCol + 1].NumberFormat = "0" + dau + "00";
            _worksheet.Cells[nRow + offset, nCol + 2].NumberFormat = "0" + dau + "00";
            _worksheet.Cells[nRow + offset, nCol + 3].NumberFormat = "0" + dau + "00";
            _worksheet.Cells[nRow + offset, nCol + 4].NumberFormat = "0" + dau + "00";
            _worksheet.Cells[nRow + offset, nCol + 5].NumberFormat = "0" + dau + "00";
            _worksheet.Cells[nRow + offset, nCol].NumberFormat = "0" + dau + "00";

        }
    }
    public static class Excelcontrunhanchia
    {
        public static void AddValueToSelectedCell(Excel.Application excelApp, string valueToAdd)
        {
            if (!double.TryParse(valueToAdd, out double value))
            {
                MessageBox.Show("Giá trị cần cộng không phải là số hợp lệ.");
                return;
            }
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;
            if (selectedRange == null)
            {
                MessageBox.Show("Không có ô nào được chọn.");
                return;
            }

            bool hasNonNumeric = false;
            using (new ExcelWriteContext(excelApp))
            using (var transaction = new ExcelBatchWriteTransaction())
            {
                Excel.Areas areas = null;
                try
                {
                    areas = selectedRange.Areas;
                    for (int areaIndex = 1; areaIndex <= areas.Count; areaIndex++)
                    {
                        Excel.Range area = null;
                        try
                        {
                            area = areas.Item[areaIndex];
                            int rows = area.Rows.Count;
                            int columns = area.Columns.Count;
                            object[,] values = TinhToan.ToMatrix(area.Value2, rows, columns);
                            object[,] formulas = TinhToan.ToMatrix(area.Formula, rows, columns);
                            for (int row = 0; row < rows; row++)
                            {
                                for (int column = 0; column < columns; column++)
                                {
                                    object raw = values[row, column];
                                    if (raw != null && double.TryParse(raw.ToString(), out double number))
                                        formulas[row, column] = number + value;
                                    else if (raw != null)
                                        hasNonNumeric = true;
                                }
                            }
                            transaction.WriteFormula(area, formulas);
                        }
                        finally
                        {
                            TinhToan.ReleaseCom(area);
                        }
                    }
                }
                finally
                {
                    TinhToan.ReleaseCom(areas);
                }
                transaction.Commit();
            }
            if (hasNonNumeric)
                MessageBox.Show("Một số ô không chứa dữ liệu số nên được giữ nguyên.");
        }
    }
    public class ExcelDongia
    {
        private readonly Excel.Worksheet _worksheet;
        private readonly Random _random = new Random();

        public ExcelDongia(Excel.Worksheet worksheet)
        {
            _worksheet = worksheet;
        }
        public void ProcessCellDonGia(int nRow, int nCol, string defaultValue, string ColumnSheetTT)
        {
         
           if (FindValueInColumn (defaultValue, nCol))
            {
                string cellValue;
                Excel.Range lastCell = _worksheet.Cells[_worksheet.Rows.Count, ColumnSheetTT].End(Excel.XlDirection.xlUp);

                // Kiểm tra lastCell có phải là null không
                if (lastCell == null || lastCell.Row < 1)
                {
                    MessageBox.Show("Không tìm thấy dữ liệu trong cột " + ColumnSheetTT, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int lastRow = lastCell.Row;
                int lastCol = lastCell.Column;
                int currentRow = 1;
                int nColNew = lastCol;  // Initialize nColNew with a valid column index

                while (currentRow <= nRow)
                {
                    cellValue = _worksheet.Cells[currentRow, nCol].Value?.ToString() ?? "";  // Ensure cellValue is never null

                    if (cellValue == defaultValue)
                    {
                        int.TryParse(_worksheet.Cells[currentRow + 1, nCol].Value?.ToString() ?? "", out nColNew);
                        currentRow += 2;
                        continue;
                    }

                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        if (nColNew >= 1 && nColNew <= _worksheet.Columns.Count)  // Ensure nColNew is within valid range
                        {
                            _worksheet.Cells[currentRow, nCol].Formula = "=" + _worksheet.Cells[currentRow, lastCol + nColNew].Address;
                            for (int j = 1; j <= 8; j++)
                            {
                                _worksheet.Cells[currentRow, nCol + j].Interior.ColorIndex = Excel.Constants.xlNone;
                                _worksheet.Cells[currentRow, nCol + j].Font.Bold = false;
                            }
                            _worksheet.Cells[currentRow, nCol + nColNew].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            _worksheet.Cells[currentRow, nCol + nColNew].Font.Bold = true;


                        }
                    }

                    currentRow++;
                }


            }    
            
        }
        private Boolean FindValueInColumn(string defaultValue, int nCol)
        {
            // Chuyển đổi cột từ chữ cái sang số (A = 1, B = 2, ... Z = 26, AA = 27, ...)
 

            Excel.Range columnRange = _worksheet.Columns[nCol];
            Excel.Range foundCell = columnRange.Find(
                What: defaultValue,
                LookIn: Excel.XlFindLookIn.xlValues,
                LookAt: Excel.XlLookAt.xlWhole
            );

            if (foundCell == null)
            {
                MessageBox.Show("Sai cột, chọn lại ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else
            {
                return true;
            }
        }


    }
}
