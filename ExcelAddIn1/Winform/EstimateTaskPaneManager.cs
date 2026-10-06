using Microsoft.Office.Core;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Winform
{
    internal static class EstimateTaskPaneManager
    {
        private sealed class PaneEntry
        {
            internal Excel.Workbook Workbook;
            internal CustomTaskPane Pane;
            internal EstimateTaskPaneControl Control;
        }

        private static readonly Dictionary<string, PaneEntry> Entries =
            new Dictionary<string, PaneEntry>(StringComparer.Ordinal);

        internal static void Show(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            string key = WorkbookKey(workbook);
            PaneEntry existing;
            if (Entries.TryGetValue(key, out existing))
            {
                if (existing.Pane != null)
                {
                    existing.Control?.RefreshOverview();
                    existing.Pane.Visible = true;
                    return;
                }
                Entries.Remove(key);
            }

            var control = new EstimateTaskPaneControl(workbook);
            CustomTaskPane pane = null;
            Excel.Window activeWindow = null;
            try
            {
                activeWindow = workbook.Application.ActiveWindow;
                pane = Globals.ThisAddIn.CustomTaskPanes.Add(
                    control,
                    "Trợ lý Dự toán",
                    activeWindow);
                pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = 430;
                pane.Visible = true;

                Entries[key] = new PaneEntry
                {
                    Workbook = workbook,
                    Pane = pane,
                    Control = control
                };
            }
            catch
            {
                if (pane != null)
                {
                    try
                    {
                        Globals.ThisAddIn.CustomTaskPanes.Remove(pane);
                    }
                    catch
                    {
                    }
                }
                control.Dispose();
                throw;
            }
            finally
            {
                if (activeWindow != null && Marshal.IsComObject(activeWindow))
                    Marshal.ReleaseComObject(activeWindow);
            }
        }

        internal static void ShowSettings(
            Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            Show(workbook);

            PaneEntry entry;
            if (Entries.TryGetValue(
                WorkbookKey(workbook),
                out entry))
            {
                entry.Control?.ShowSettings();
                if (entry.Pane != null)
                    entry.Pane.Visible = true;
            }
        }

        internal static void Refresh(Excel.Workbook workbook)
        {
            if (workbook == null)
                return;
            PaneEntry entry;
            if (Entries.TryGetValue(WorkbookKey(workbook), out entry))
                entry.Control?.RefreshOverview();
        }

        internal static void ShowPackages(Excel.Workbook workbook)
        {
            Show(workbook);
            Entries[WorkbookKey(workbook)].Control.ShowPackages();
        }

        internal static void ShowReports(Excel.Workbook workbook)
        {
            Show(workbook);
            Entries[WorkbookKey(workbook)].Control.ShowReports();
        }

        internal static void CloseForWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
                return;
            string key;
            try
            {
                key = WorkbookKey(workbook);
            }
            catch
            {
                return;
            }

            PaneEntry entry;
            if (!Entries.TryGetValue(key, out entry))
                return;
            Entries.Remove(key);
            if (entry.Pane != null)
            {
                try
                {
                    Globals.ThisAddIn.CustomTaskPanes.Remove(entry.Pane);
                }
                catch
                {
                    try
                    {
                        entry.Pane.Visible = false;
                    }
                    catch
                    {
                    }
                }
            }
            entry.Control?.Dispose();
        }

        internal static void CloseAll()
        {
            foreach (PaneEntry entry in new List<PaneEntry>(Entries.Values))
            {
                if (entry.Pane != null)
                {
                    try
                    {
                        Globals.ThisAddIn.CustomTaskPanes.Remove(entry.Pane);
                    }
                    catch
                    {
                    }
                }
                entry.Control?.Dispose();
            }
            Entries.Clear();
        }

        private static string WorkbookKey(Excel.Workbook workbook)
        {
            IntPtr unknown = IntPtr.Zero;
            try
            {
                unknown = Marshal.GetIUnknownForObject(workbook);
                return unknown.ToInt64().ToString("X16", System.Globalization.CultureInfo.InvariantCulture);
            }
            finally
            {
                if (unknown != IntPtr.Zero)
                    Marshal.Release(unknown);
            }
        }
    }
}
