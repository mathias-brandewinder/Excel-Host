using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService
{
    public class ExcelService : IExcelService
    {
        public void DoStuff(string message)
        {
            Debug.WriteLine(String.Format("Hi: {0}", message));
        }
    }
}