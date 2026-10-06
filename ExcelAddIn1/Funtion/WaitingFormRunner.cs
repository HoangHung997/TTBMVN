using ExcelAddIn1;
using ExcelAddIn1.Funtion;
using System;
using System.Threading;
using System.Windows.Forms;

public static class DaoDatRunner
{
    private static Thread _uiThread;


    public static void ShowWaitingForm()
    {
        _uiThread = new Thread(() =>
        {
            SplashManager.ShowForm("Vui lòng chờ!", "Đang xử lý dữ liệu...");
        });

        _uiThread.SetApartmentState(ApartmentState.STA);
        _uiThread.Start();
    }

    public static void CloseWaitingForm()
    {
        SplashManager.CloseForm();
    }
}
