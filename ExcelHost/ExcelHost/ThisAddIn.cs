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
    using System.ServiceModel;

    public partial class ThisAddIn
    {
        public ServiceHost Host { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var workbooks = excel.Workbooks;
            var workbook = workbooks.Add();



            ExcelService.ExcelService.Workbook = workbook;

            var host = new ServiceHost(typeof(ExcelService.ExcelService),
                                       new[] { new Uri("net.pipe://localhost") });

            host.AddServiceEndpoint(typeof(ExcelService.IExcelService), new NetNamedPipeBinding(), "excel");
            host.Open();
            
            this.Host = host;

            var fsiPath = @"C:\Program Files (x86)\Microsoft F#\v4.0\fsi.exe";

            var info = new ProcessStartInfo();
            var fsiProcess = new Process();
            //info.RedirectStandardInput = true
            //info.RedirectStandardOutput = true
            //info.UseShellExecute = false
            //info.CreateNoWindow = true
            info.FileName = fsiPath;

            fsiProcess.StartInfo = info;
            fsiProcess.Start();

            //var form = new TheForm();
            //form.Show();
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
