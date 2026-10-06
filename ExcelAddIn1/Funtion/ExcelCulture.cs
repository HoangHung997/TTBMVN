using System;
using System.Globalization;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Funtion
{
    public static class ExcelCulture
    {
        public static CultureInfo GetNumberCulture(Excel.Application application)
        {
            CultureInfo culture = (CultureInfo)CultureInfo.CurrentCulture.Clone();

            try
            {
                string decimalSeparator = Convert.ToString(application.International[Excel.XlApplicationInternational.xlDecimalSeparator]);
                string thousandsSeparator = Convert.ToString(application.International[Excel.XlApplicationInternational.xlThousandsSeparator]);

                if (!string.IsNullOrEmpty(decimalSeparator))
                    culture.NumberFormat.NumberDecimalSeparator = decimalSeparator;

                if (!string.IsNullOrEmpty(thousandsSeparator))
                    culture.NumberFormat.NumberGroupSeparator = thousandsSeparator;
            }
            catch
            {
                // Keep current culture if Excel separators cannot be read yet.
            }

            return culture;
        }

        public static void ApplyToCurrentThread(Excel.Application application)
        {
            CultureInfo culture = GetNumberCulture(application);
            Thread.CurrentThread.CurrentCulture = culture;
            CultureInfo.DefaultThreadCurrentCulture = culture;
        }
    }
}
