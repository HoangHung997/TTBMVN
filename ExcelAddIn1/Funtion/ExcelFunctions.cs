using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public class ExcelFunctions
    {
     
        public static int FindRow(string searchString, Excel.Range columnRange)
        {
            // Duyệt qua từng ô trong cột
            foreach (Excel.Range cell in columnRange)
            {
                if (cell.Value2 != null && cell.Value2.ToString() == searchString)
                {
                    // Trả về số dòng nếu tìm thấy
                    return cell.Row;
                }
            }

            // Nếu không tìm thấy, trả về -1
            return -1;
        }
        public static (int StartRow, int EndRow) FindRowlist(string searchString, Excel.Range columnRange)
        {
            Excel.Range startCell = null;
            Excel.Range endCell = null;

            // Tìm ô bắt đầu (Start + searchString)
            foreach (Excel.Range cell in columnRange)
            {
                if (cell.Value2 != null && cell.Value2.ToString() == "Start" + searchString)
                {
                    startCell = cell;
                    break;
                }
            }

            // Nếu không tìm thấy ô bắt đầu
            if (startCell == null)
            {
                throw new InvalidOperationException("Not Found (Start)");
            }

            // Tìm ô kết thúc (End + searchString)
            foreach (Excel.Range cell in columnRange)
            {
                if (cell.Value2 != null && cell.Value2.ToString() == "End" + searchString)
                {
                    endCell = cell;
                    break;
                }
            }

            // Nếu không tìm thấy ô kết thúc
            if (endCell == null)
            {
                throw new InvalidOperationException("Not Found (End)");
            }

            // Kiểm tra nếu không có dòng nào giữa Start và End
            if (startCell.Row + 1 > endCell.Row)
            {
                throw new InvalidOperationException("No Range Between Start and End");
            }

            // Trả về cặp dòng bắt đầu và dòng kết thúc
            return (startCell.Row, endCell.Row);
        }

    }
}
