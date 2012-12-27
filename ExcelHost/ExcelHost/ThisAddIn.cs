using Office = Microsoft.Office.Core;

namespace ExcelHost
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var workbooks = excel.Workbooks;
            var workbook = workbooks.Add();

            ExcelService.Workbook = workbook;
            ExcelService.Start();

            //var fsiPath = @"C:\Program Files (x86)\Microsoft F#\v4.0\fsi.exe";

            //var info = new ProcessStartInfo();
            //var fsiProcess = new Process();
            //info.RedirectStandardInput = true
            //info.RedirectStandardOutput = true
            //info.UseShellExecute = false
            //info.CreateNoWindow = true
            //info.FileName = fsiPath;

            //fsiProcess.StartInfo = info;
            //fsiProcess.Start();

            var form = new TheForm();
            form.Show();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            ExcelService.Stop();
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
