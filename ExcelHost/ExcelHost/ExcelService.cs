using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService
{
    using Microsoft.Office.Interop.Excel;

    public class ExcelService : IExcelService
    {
        internal static Workbook Workbook { get; set; }

        public void DoStuff(string message)
        {
            var worksheet = Workbook.ActiveSheet;
            worksheet.Cells[1, 1] = message;
        }

        public string GrabIt()
        {
            var worksheet = Workbook.ActiveSheet;
            var content = worksheet.Cells[1, 1].Value2;
            return content ?? "N/A";

            //return "Hogmonaut";
        }
    }
}