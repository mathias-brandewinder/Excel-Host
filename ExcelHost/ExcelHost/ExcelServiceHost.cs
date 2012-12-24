using System;
using System.ServiceModel;

namespace ExcelHost
{
    public class ExcelServiceHost
    {
        public void Run()
        {
            using (var host = new ServiceHost(typeof (ExcelService.ExcelService),
                                              new[] {new Uri("net.pipe://localhost")}))
            {
                host.AddServiceEndpoint(typeof (ExcelService.IExcelService), new NetNamedPipeBinding(), "excel");
                host.Open();
            }
        }
    }
}