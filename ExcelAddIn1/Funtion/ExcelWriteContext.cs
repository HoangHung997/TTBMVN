using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public sealed class ExcelWriteContext : IDisposable
    {
        private readonly Excel.Application application;
        private readonly bool oldScreenUpdating;
        private readonly Excel.XlCalculation oldCalculation;
        private readonly bool oldEnableEvents;
        private readonly bool oldDisplayAlerts;
        private bool disposed;

        public ExcelWriteContext(Excel.Application application, bool displayAlerts = false)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            oldScreenUpdating = application.ScreenUpdating;
            oldCalculation = application.Calculation;
            oldEnableEvents = application.EnableEvents;
            oldDisplayAlerts = application.DisplayAlerts;

            try
            {
                application.ScreenUpdating = false;
                application.Calculation = Excel.XlCalculation.xlCalculationManual;
                application.EnableEvents = false;
                application.DisplayAlerts = displayAlerts;
            }
            catch (Exception setupException)
            {
                try
                {
                    RestoreCore();
                }
                catch (Exception restoreException)
                {
                    throw new AggregateException(
                        "Khong thiet lap duoc Excel write context va khong phuc hoi day du trang thai.",
                        setupException,
                        restoreException);
                }
                throw;
            }
        }

        public void Dispose()
        {
            if (disposed)
                return;

            try
            {
                RestoreCore();
            }
            catch (Exception ex)
            {
                RuntimeLogger.Log(ex, "Restore Excel application state");
            }
            finally
            {
                disposed = true;
            }
        }

        private void RestoreCore()
        {
            application.Calculation = oldCalculation;
            application.EnableEvents = oldEnableEvents;
            application.DisplayAlerts = oldDisplayAlerts;
            application.ScreenUpdating = oldScreenUpdating;
        }
    }
}
