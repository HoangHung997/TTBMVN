using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class ExcelBatchWriteTransaction : IDisposable
    {
        private readonly List<Snapshot> snapshots = new List<Snapshot>();
        private bool committed;
        private bool disposed;

        public int BatchCount => snapshots.Count;
        public long CellCount { get; private set; }

        public void WriteFormula(Excel.Range range, object value)
        {
            Write(range, value, true);
        }

        public void WriteValue2(Excel.Range range, object value)
        {
            Write(range, value, false);
        }

        public void Commit()
        {
            ThrowIfDisposed();
            committed = true;
        }

        public void Rollback()
        {
            ThrowIfDisposed();
            Restore();
            committed = true;
        }

        public void Dispose()
        {
            if (disposed)
                return;
            try
            {
                if (!committed)
                    Restore();
            }
            finally
            {
                foreach (Snapshot snapshot in snapshots)
                    snapshot.Dispose();
                snapshots.Clear();
                disposed = true;
            }
        }

        private void Write(Excel.Range range, object value, bool formula)
        {
            ThrowIfDisposed();
            if (committed)
                throw new InvalidOperationException("Transaction da commit.");
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            snapshots.Add(new Snapshot(range, range.Formula));
            CellCount = checked(CellCount + Convert.ToInt64(range.CountLarge));
            if (formula)
                range.Formula = value;
            else
                range.Value2 = value;
        }

        private void Restore()
        {
            Exception first = null;
            for (int index = snapshots.Count - 1; index >= 0; index--)
            {
                try
                {
                    snapshots[index].Restore();
                }
                catch (Exception ex)
                {
                    first = first ?? ex;
                    RuntimeLogger.Log(ex, "Rollback Excel batch write");
                }
            }
            if (first != null)
                throw new InvalidOperationException("Khong the rollback day du cac batch Excel.", first);
        }

        private void ThrowIfDisposed()
        {
            if (disposed)
                throw new ObjectDisposedException(nameof(ExcelBatchWriteTransaction));
        }

        private sealed class Snapshot : IDisposable
        {
            public Snapshot(Excel.Range range, object formula)
            {
                Worksheet = range.Worksheet as Excel.Worksheet;
                Address = range.get_Address(
                    true,
                    true,
                    Excel.XlReferenceStyle.xlA1,
                    Type.Missing,
                    Type.Missing);
                Formula = formula;
            }

            public Excel.Worksheet Worksheet { get; private set; }
            public string Address { get; }
            public object Formula { get; }

            public void Restore()
            {
                Excel.Range range = null;
                try
                {
                    range = Worksheet.Range[Address];
                    range.Formula = Formula;
                }
                finally
                {
                    Release(range);
                }
            }

            public void Dispose()
            {
                Release(Worksheet);
                Worksheet = null;
            }

            private static void Release(object value)
            {
                if (value != null && Marshal.IsComObject(value))
                    Marshal.ReleaseComObject(value);
            }
        }
    }
}
