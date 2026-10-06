using ExcelAddIn1.Winform;
using System;

using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace ExcelAddIn1.Funtion
{
    public static class SplashManager   // ✅ thêm class bao ngoài
    {

        private static Waiting splashForm;
        private static Thread splashThread;
        private static readonly object locker = new object();

        /// <summary>
        /// Hiển thị form chờ (giống SplashScreenManager.ShowForm)
        /// </summary>
        public static void ShowForm(string caption = "Vui lòng chờ!", string description = "Đang xử lý dữ liệu...")
        {
            lock (locker)
            {
                if (splashForm != null) return;

                splashThread = new Thread(() =>
                {
                    splashForm = new Waiting(caption, description);
                    Application.Run(splashForm);
                });

                splashThread.IsBackground = true;
                splashThread.SetApartmentState(ApartmentState.STA);
                splashThread.Start();

                // Đợi cho đến khi form hiển thị
                while (splashForm == null || !splashForm.IsHandleCreated)
                    Thread.Sleep(10);
            }
        }

        /// <summary>
        /// Đóng form chờ (giống SplashScreenManager.CloseForm)
        /// </summary>
        public static void CloseForm()
        {
            lock (locker)
            {
                if (splashForm == null) return;

                try
                {
                    if (splashForm.InvokeRequired)
                    {
                        splashForm.Invoke(new Action(() => splashForm.Close()));
                    }
                    else
                    {
                        splashForm.Close();
                    }
                }
                catch { /* bỏ qua lỗi khi đóng */ }

                splashForm = null;
                splashThread = null;
            }
        }

        /// <summary>
        /// Cập nhật dòng mô tả (nếu cần)
        /// </summary>
        public static void SetDescription(string text)
        {
            if (splashForm == null) return;

            try
            {
                if (splashForm.InvokeRequired)
                    splashForm.Invoke(new Action(() => splashForm.UpdateDescription(text)));
                else
                    splashForm.UpdateDescription(text);
            }
            catch { }
        }
    }
}
