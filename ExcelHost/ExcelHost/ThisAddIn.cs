using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Diagnostics;

namespace ExcelHost
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Interop.Excel.Worksheet TargetSheet { get; set; }

        public void DoSomethingWithMessage(string message)
        {
            this.TargetSheet.Cells[1, 1].Value2 = message;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var workbooks = excel.Workbooks;
            var workbook = workbooks.Add();
            this.TargetSheet = workbook.ActiveSheet; ;

            this.DoSomethingWithMessage("Hello, World");

            var fsiPath = @"C:\Program Files (x86)\Microsoft F#\v4.0\fsi.exe";

            var info = new ProcessStartInfo();
            var fsiProcess = new Process();
            var appDomain = AppDomain.CurrentDomain;
            info.Domain = appDomain.FriendlyName;
            //info.RedirectStandardInput = true
            //info.RedirectStandardOutput = true
            //info.UseShellExecute = false
            //info.CreateNoWindow = true
            info.FileName = fsiPath;

            var myService = new ExcelServiceHost();
            myService.Run();

            fsiProcess.StartInfo = info;
            fsiProcess.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
