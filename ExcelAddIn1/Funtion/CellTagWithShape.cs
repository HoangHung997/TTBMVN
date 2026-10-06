using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn1.Funtion
{
    public static class CellTagWithShape
    {
        private const string TagPrefix = "TAG__";        // để lọc nhanh trong Shapes
        private const string AltPrefix = "ALT__TAG__";   // lưu tag ở AltText

        /// Gắn “tag” vào 1 ô bằng cách chèn 1 shape vô hình neo đúng ô đó
        public static void AddTag(string tag, Excel.Range rng)
        {
            var ws = rng.Worksheet;
            var app = ws.Application;

            // Nếu đã có thì xóa trước
            RemoveTag(tag, ws);

            // Nếu là ô gộp, neo vào vùng gộp
            Excel.Range anchor = rng.MergeCells ? rng.MergeArea.Cells[1, 1] : rng.Cells[1, 1];

            // Change this line:
            // var shp = ws.Shapes.AddShape(
            //     (int)Office.MsoAutoShapeType.msoShapeRectangle,
            //     (float)anchor.Left, (float)anchor.Top,
            //     (float)anchor.Width, (float)anchor.Height);

            // To this:
            var shp = ws.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                (float)anchor.Left, (float)anchor.Top,
                (float)anchor.Width, (float)anchor.Height);

            // Đặt tên/metadata để truy xuất
            string shapeName = TagPrefix + tag;
            shp.Name = shapeName;
            shp.AlternativeText = AltPrefix + tag;

            // Ẩn hoàn toàn
            shp.Line.Visible = Office.MsoTriState.msoFalse;
            shp.Fill.Visible = Office.MsoTriState.msoFalse;
            shp.Visible = Office.MsoTriState.msoFalse;           // không thể click/xóa tình cờ
            shp.Placement = Excel.XlPlacement.xlMoveAndSize;     // trôi theo ô khi chèn/xóa
            /*shp.PrintObject = Office.MsoTriState.msoFalse;   */    // không in
            shp.LockAspectRatio = Office.MsoTriState.msoTrue;
            // (Không cần Locked vì bạn không bật Protect sheet)

            // Nếu muốn tag cả vùng bảng, có thể set Width/Height = anchor.MergeArea.Width/Height
        }

        /// Lấy lại Range đang được “tag” (top-left cell của shape)
        /// Lấy lại Range đang được “tag” (top-left cell của shape)
        public static Excel.Range GetTaggedRange(string tag, Excel.Worksheet ws = null)
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            // Nếu truyền sheet cụ thể thì chỉ tìm trong sheet đó
            if (ws != null)
            {
                string shapeName = TagPrefix + tag;
                string altText = AltPrefix + tag;

                foreach (Excel.Shape s in ws.Shapes)
                {
                    if (s.Name == shapeName || s.AlternativeText == altText)
                        return s.TopLeftCell;
                }
                return null;
            }

            // Nếu ws == null thì duyệt toàn bộ workbook
            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                string shapeName = TagPrefix + tag;
                string altText = AltPrefix + tag;

                foreach (Excel.Shape s in sheet.Shapes)
                {
                    if (s.Name == shapeName || s.AlternativeText == altText)
                        return s.TopLeftCell;
                }
            }

            return null; // không tìm thấy
        }


        /// Xóa tag (nếu cần)
        public static void RemoveTag(string tag, Excel.Worksheet ws = null)
        {
            if (ws == null)
                ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string shapeName = TagPrefix + tag;
            string altText = AltPrefix + tag;

            // Xóa tất cả shape trùng tag
            var toDelete = new List<Excel.Shape>();
            foreach (Excel.Shape s in ws.Shapes)
            {
                if (s.Name == shapeName || s.AlternativeText == altText)
                    toDelete.Add(s);
            }
            foreach (var s in toDelete) s.Delete();
        }

        /// Ví dụ: gắn tag cho ô đang chọn
        public static void TagSelectionAs(string tag)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Excel.Range;
            if (sel == null) return;
            AddTag(tag, sel);
        }
        public static void MoveTag(string tag, Excel.Range newRng, Excel.Worksheet ws = null)
        {
            if (ws == null)
                ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            var shp = FindTagShape(tag, ws);
            if (shp == null) { AddTag(tag, newRng); return; }

            Excel.Range anchor = newRng.MergeCells ? newRng.MergeArea.Cells[1, 1] : newRng.Cells[1, 1];

            shp.Left = (float)anchor.Left;
            shp.Top = (float)anchor.Top;
            shp.Width = (float)anchor.Width;
            shp.Height = (float)anchor.Height;
            shp.Placement = Excel.XlPlacement.xlMoveAndSize;
        }
        /// Ví dụ: đọc giá trị từ ô đã tag
        public static object ReadValue(string tag)
        {
            var rng = GetTaggedRange(tag);
            return rng?.Value2; // null nếu chưa tag/shape bị xóa
        }

        /// Ví dụ: nhảy con trỏ tới ô đã tag
        public static void GoToTag(string tag)
        {
            var rng = GetTaggedRange(tag);
            if (rng != null) rng.Select();
        }
        private static Excel.Shape FindTagShape(string tag, Excel.Worksheet ws)
        {
            string name = TagPrefix + tag;
            string alt = AltPrefix + tag;

            foreach (Excel.Shape s in ws.Shapes)
                if (s.Name == name || s.AlternativeText == alt)
                    return s;
            return null;
        }
    }
}
