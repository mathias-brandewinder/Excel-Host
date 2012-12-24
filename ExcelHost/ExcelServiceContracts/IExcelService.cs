using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService
{
    [ServiceContract]
    public interface IExcelService
    {
        [OperationContract]
        void DoStuff(string message);
    }
}