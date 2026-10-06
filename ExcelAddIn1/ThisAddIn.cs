using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelAddIn1.Funtion;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        internal WorkbookSheetChangeCoordinator SheetChangeCoordinator { get; private set; }

        public void Undo()
        {
            ExcelUndoManager.Undo();
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ExcelCulture.ApplyToCurrentThread(Application);
            RuntimeLogger.LogMessage("Add-in startup", new System.Collections.Generic.Dictionary<string, string>
            {
                ["Product"] = AppInfo.ProductName,
                ["Version"] = AppInfo.Version,
                ["License"] = LicenseManager.GetStatus().Message
            });

            SheetChangeCoordinator = new WorkbookSheetChangeCoordinator(Application);
            Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;

        }





        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            if (!cancel)
                ExcelAddIn1.Winform.EstimateTaskPaneManager.CloseForWorkbook(workbook);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            }
            catch
            {
            }
            ExcelAddIn1.Winform.EstimateTaskPaneManager.CloseAll();
            SheetChangeCoordinator?.Dispose();
            SheetChangeCoordinator = null;
        }
  
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

    }

}
