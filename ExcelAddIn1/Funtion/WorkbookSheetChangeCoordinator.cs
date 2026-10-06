using ExcelAddIn1.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public interface IWorkbookSheetListObserver
    {
        void OnWorkbookSheetsChanged(IReadOnlyList<WorkbookSheetDescriptor> sheets);
    }

    public sealed class WorkbookSheetChangeCoordinator : IDisposable
    {
        private sealed class Subscription
        {
            public Excel.Workbook Workbook;
            public WeakReference Observer;
            public string LastSignature;
            public bool ForceRefresh;
            public int ConsecutiveFailures;
        }

        private sealed class SubscriptionToken : IDisposable
        {
            private WorkbookSheetChangeCoordinator owner;
            private Subscription subscription;

            public SubscriptionToken(WorkbookSheetChangeCoordinator owner, Subscription subscription)
            {
                this.owner = owner;
                this.subscription = subscription;
            }

            public void Dispose()
            {
                WorkbookSheetChangeCoordinator currentOwner = owner;
                Subscription currentSubscription = subscription;
                owner = null;
                subscription = null;
                currentOwner?.Unsubscribe(currentSubscription);
            }
        }

        private readonly Excel.Application application;
        private readonly Timer timer;
        private readonly List<Subscription> subscriptions = new List<Subscription>();
        private bool disposed;
        private bool polling;

        public event Action<Excel.Workbook> WorkbookClosing;

        public WorkbookSheetChangeCoordinator(Excel.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            timer = new Timer { Interval = 350 };
            timer.Tick += Timer_Tick;

            application.WorkbookActivate += Application_WorkbookActivate;
            application.WorkbookOpen += Application_WorkbookOpen;
            application.WorkbookNewSheet += Application_WorkbookNewSheet;
            application.SheetActivate += Application_SheetActivate;
            application.SheetBeforeDelete += Application_SheetBeforeDelete;
            application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
        }

        public IDisposable Subscribe(Excel.Workbook workbook, IWorkbookSheetListObserver observer)
        {
            if (disposed)
                throw new ObjectDisposedException(nameof(WorkbookSheetChangeCoordinator));
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (observer == null)
                throw new ArgumentNullException(nameof(observer));

            var subscription = new Subscription
            {
                Workbook = workbook,
                Observer = new WeakReference(observer),
                ForceRefresh = true
            };
            subscriptions.Add(subscription);
            timer.Start();
            PollNow();
            return new SubscriptionToken(this, subscription);
        }

        public void PollNow()
        {
            if (disposed || polling)
                return;

            polling = true;
            try
            {
                for (int index = subscriptions.Count - 1; index >= 0; index--)
                {
                    Subscription subscription = subscriptions[index];
                    var observer = subscription.Observer.Target as IWorkbookSheetListObserver;
                    if (observer == null)
                    {
                        subscriptions.RemoveAt(index);
                        continue;
                    }

                    try
                    {
                        IReadOnlyList<WorkbookSheetDescriptor> sheets = CaptureSnapshot(subscription.Workbook);
                        string signature = WorkbookSheetList.BuildSignature(sheets);
                        if (!subscription.ForceRefresh &&
                            string.Equals(signature, subscription.LastSignature, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        subscription.LastSignature = signature;
                        subscription.ForceRefresh = false;
                        subscription.ConsecutiveFailures = 0;
                        observer.OnWorkbookSheetsChanged(sheets);
                    }
                    catch (COMException ex)
                    {
                        RuntimeLogger.Log(ex, "Poll workbook sheet list");
                        subscription.ConsecutiveFailures++;
                        if (subscription.ConsecutiveFailures >= 3)
                            subscriptions.RemoveAt(index);
                    }
                    catch (InvalidComObjectException ex)
                    {
                        RuntimeLogger.Log(ex, "Poll closed workbook sheet list");
                        subscriptions.RemoveAt(index);
                    }
                }

                if (subscriptions.Count == 0)
                    timer.Stop();
            }
            finally
            {
                polling = false;
            }
        }

        public static IReadOnlyList<WorkbookSheetDescriptor> CaptureSnapshot(Excel.Workbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var result = new List<WorkbookSheetDescriptor>();
            Excel.Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                for (int index = 1; index <= worksheets.Count; index++)
                {
                    Excel.Worksheet worksheet = null;
                    try
                    {
                        worksheet = worksheets.Item[index] as Excel.Worksheet;
                        if (worksheet == null)
                            continue;

                        string name = worksheet.Name;
                        string key;
                        try
                        {
                            key = worksheet.CodeName;
                        }
                        catch (COMException)
                        {
                            key = name;
                        }

                        if (string.IsNullOrWhiteSpace(key))
                            key = GetRuntimeWorksheetKey(worksheet);
                        result.Add(new WorkbookSheetDescriptor(key, name));
                    }
                    finally
                    {
                        ReleaseComObject(worksheet);
                    }
                }
            }
            finally
            {
                ReleaseComObject(worksheets);
            }

            return result.AsReadOnly();
        }

        public void Dispose()
        {
            if (disposed)
                return;

            disposed = true;
            timer.Stop();
            timer.Tick -= Timer_Tick;
            timer.Dispose();

            application.WorkbookActivate -= Application_WorkbookActivate;
            application.WorkbookOpen -= Application_WorkbookOpen;
            application.WorkbookNewSheet -= Application_WorkbookNewSheet;
            application.SheetActivate -= Application_SheetActivate;
            application.SheetBeforeDelete -= Application_SheetBeforeDelete;
            application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            WorkbookClosing = null;
            subscriptions.Clear();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            PollNow();
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            MarkForRefresh(workbook);
        }

        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            MarkForRefresh(workbook);
        }

        private void Application_WorkbookNewSheet(Excel.Workbook workbook, object sheet)
        {
            MarkForRefresh(workbook);
        }

        private void Application_SheetActivate(object sheet)
        {
            Excel.Worksheet worksheet = sheet as Excel.Worksheet;
            if (worksheet == null)
                return;

            Excel.Workbook workbook = null;
            try
            {
                workbook = worksheet.Parent as Excel.Workbook;
                MarkForRefresh(workbook);
            }
            finally
            {
                ReleaseComObject(workbook);
            }
        }

        private void Application_SheetBeforeDelete(object sheet)
        {
            Excel.Worksheet worksheet = sheet as Excel.Worksheet;
            if (worksheet == null)
                return;

            Excel.Workbook workbook = null;
            try
            {
                workbook = worksheet.Parent as Excel.Workbook;
                MarkForRefresh(workbook);
            }
            finally
            {
                ReleaseComObject(workbook);
            }
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            if (!cancel)
                WorkbookClosing?.Invoke(workbook);
            MarkForRefresh(workbook);
        }

        private void MarkForRefresh(Excel.Workbook workbook)
        {
            if (workbook == null)
                return;

            foreach (Subscription subscription in subscriptions)
            {
                if (IsSameWorkbook(subscription.Workbook, workbook))
                    subscription.ForceRefresh = true;
            }
        }

        private void Unsubscribe(Subscription subscription)
        {
            if (subscription == null)
                return;

            subscriptions.Remove(subscription);
            if (subscriptions.Count == 0)
                timer.Stop();
        }

        private static bool IsSameWorkbook(Excel.Workbook left, Excel.Workbook right)
        {
            if (left == null || right == null)
                return false;
            if (ReferenceEquals(left, right))
                return true;

            try
            {
                return string.Equals(left.FullName, right.FullName, StringComparison.OrdinalIgnoreCase);
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static string GetRuntimeWorksheetKey(Excel.Worksheet worksheet)
        {
            IntPtr unknown = IntPtr.Zero;
            try
            {
                unknown = Marshal.GetIUnknownForObject(worksheet);
                return "COM:" + unknown.ToInt64().ToString("X16", CultureInfo.InvariantCulture);
            }
            finally
            {
                if (unknown != IntPtr.Zero)
                    Marshal.Release(unknown);
            }
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
    }
}
