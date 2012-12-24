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
    }
}