using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFsiService
{
    [ServiceContract]
    public interface IFsiOperationContract
    {
        [OperationContract]
        void HelloWorld(string message);
    }

    [ServiceContract]
    public class FsiOperationContract : IFsiOperationContract
    {
        [OperationContract]
        public void HelloWorld(string message)
        {
            Debug.WriteLine("Hi: {0}", message);
        }
    }
}