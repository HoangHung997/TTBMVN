using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    [Flags]
    public enum CellProperties
    {
        None = 0,
        Value = 1,
        Formula = 2,
        NumberFormat = 4,
        FontBold = 8,
        FontItalic = 16,
        FontUnderline = 32,
        FontColor = 64,
        FontSize = 128,
        FontName = 256,
        InteriorColor = 512,
        All = ~0
    }

    public class ExcelUndoManager
    {
        public class CellState
        {
            public string Address { get; set; }

            public object Value { get; set; }
            public string Formula { get; set; }
            public string NumberFormat { get; set; }

            public bool FontBold { get; set; }
            public bool FontItalic { get; set; }
            public Excel.XlUnderlineStyle FontUnderline { get; set; }
            public int FontColor { get; set; }
            public double FontSize { get; set; }
            public string FontName { get; set; }

            public int InteriorColor { get; set; }
        }

        private static Stack<List<CellState>> undoStack = new Stack<List<CellState>>();
        private static Stack<List<CellState>> redoStack = new Stack<List<CellState>>();

        private static Excel.Application excelAppForUndoRedo;

        // Lưu trạng thái Undo, truyền Range và lựa chọn thuộc tính
        public static void SaveSnapshotForUndo(Excel.Range range, CellProperties properties = CellProperties.All)
        {
            if (range == null) throw new ArgumentNullException(nameof(range));
            excelAppForUndoRedo = range.Application;

            var snapshot = new List<CellState>();
            foreach (Excel.Range cell in range.Cells)
            {
                var state = new CellState
                {
                    Address = cell.Address[false, false],
                };

                if (properties.HasFlag(CellProperties.Value))
                    state.Value = cell.Value2;

                if (properties.HasFlag(CellProperties.Formula))
                    state.Formula = cell.Formula;

                if (properties.HasFlag(CellProperties.NumberFormat))
                    state.NumberFormat = cell.NumberFormat;

                var font = cell.Font;
                if (properties.HasFlag(CellProperties.FontBold))
                    state.FontBold = font.Bold;
                if (properties.HasFlag(CellProperties.FontItalic))
                    state.FontItalic = font.Italic;
                if (properties.HasFlag(CellProperties.FontUnderline))
                    state.FontUnderline = font.Underline != null
                    ? (XlUnderlineStyle)(int)font.Underline
                    : XlUnderlineStyle.xlUnderlineStyleNone;
                if (properties.HasFlag(CellProperties.FontColor))
                    state.FontColor = (int)font.Color;
                if (properties.HasFlag(CellProperties.FontSize))
                    state.FontSize = font.Size;
                if (properties.HasFlag(CellProperties.FontName))
                    state.FontName = font.Name;

                if (properties.HasFlag(CellProperties.InteriorColor))
                    state.InteriorColor = (int)cell.Interior.Color;

                snapshot.Add(state);
            }

            undoStack.Push(snapshot);
            redoStack.Clear();
        }

        // Undo
        public static void Undo()
        {
            if (undoStack.Count == 0) return;
            if (excelAppForUndoRedo == null) return;

            var sheet = excelAppForUndoRedo.ActiveSheet as Excel.Worksheet;
            if (sheet == null) return;

            var currentState = GetCurrentState(sheet, undoStack.Peek());

            var previousState = undoStack.Pop();

            redoStack.Push(currentState);

            ApplyState(sheet, previousState);
        }

        // Redo
        public static void Redo()
        {
            if (redoStack.Count == 0) return;
            if (excelAppForUndoRedo == null) return;

            var sheet = excelAppForUndoRedo.ActiveSheet as Excel.Worksheet;
            if (sheet == null) return;

            var currentState = GetCurrentState(sheet, redoStack.Peek());

            var redoState = redoStack.Pop();

            undoStack.Push(currentState);

            ApplyState(sheet, redoState);
        }

        // Lấy trạng thái hiện tại từ worksheet dựa trên địa chỉ snapshot
        private static List<CellState> GetCurrentState(Excel.Worksheet sheet, List<CellState> refSnapshot)
        {
            var state = new List<CellState>();
            foreach (var cellInfo in refSnapshot)
            {
                var cell = sheet.Range[cellInfo.Address];
                var s = new CellState
                {
                    Address = cell.Address[false, false],
                    Value = cell.Value2,
                    Formula = cell.Formula,
                    NumberFormat = cell.NumberFormat,
                    FontBold = cell.Font.Bold,
                    FontItalic = cell.Font.Italic,
                    FontUnderline = cell.Font.Underline,
                    FontColor = (int)cell.Font.Color,
                    FontSize = cell.Font.Size,
                    FontName = cell.Font.Name,
                    InteriorColor = (int)cell.Interior.Color
                };
                state.Add(s);
            }
            return state;
        }

        // Áp dụng trạng thái lên worksheet
        private static void ApplyState(Excel.Worksheet sheet, List<CellState> snapshot)
        {
            foreach (var state in snapshot)
            {
                var cell = sheet.Range[state.Address];

                if (state.Formula != null)
                    cell.Formula = state.Formula;
                else
                    cell.Value2 = state.Value;

                cell.NumberFormat = state.NumberFormat;

                cell.Font.Bold = state.FontBold;
                cell.Font.Italic = state.FontItalic;
                cell.Font.Underline = state.FontUnderline;
                cell.Font.Color = state.FontColor;
                cell.Font.Size = state.FontSize;
                cell.Font.Name = state.FontName;

                cell.Interior.Color = state.InteriorColor;
            }
        }

        // Đăng ký lệnh Undo cho Excel, actionName là tên hiển thị trong Undo menu Excel
        public static void RegisterUndo(string actionName, Excel.Application app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            app.OnUndo(actionName, "ThisAddIn.Undo");
            excelAppForUndoRedo = app; // đảm bảo giữ app
        }
    }
}
