using System;
using System.Diagnostics;
using System.ServiceModel;
using System.ServiceModel.Description;

namespace ExcelService
{
    using Microsoft.Office.Interop.Excel;

    internal class ExcelService : IExcelService
    {
        internal static Workbook Workbook { get; set; }
        private static ServiceHost HostedInstance { get; set; }

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
        }

        public void DoMultipleStuff(object[][] stuff)
        {
            for (var row = 0; row < stuff.GetLength(0); ++row)
            {
                for (var col = 0; col < stuff[row].GetLength(0); ++col)
                {
                    Debug.WriteLine(stuff[row][col]);
                }
            }
        }

        public object[][] GrabMultipleIt()
        {
            var fakeResult = new object[][]
                                 {
                                     new object[] { "hi", "my", "name", "is", "hogmonaut" },
                                     new object[] { 1, 2, 3, 4, 5 },
                                     new object[] { "the", "answer", "is", 42, "!" },
                                     new object[] { "lolwhut", null, null, null, null }
                                 };

            return fakeResult;
        }

        public static ServiceHost Run()
        {
            if (HostedInstance == null)
            {
                var endPointUri = new Uri("net.pipe://localhost");
                var host = new ServiceHost(typeof(ExcelService), new[] {endPointUri});

                host.AddServiceEndpoint(typeof(IExcelService), new NetNamedPipeBinding(), "excel");

                var smb = new ServiceMetadataBehavior();
                host.Description.Behaviors.Add(smb);
                host.AddServiceEndpoint(ServiceMetadataBehavior.MexContractName, MetadataExchangeBindings.CreateMexNamedPipeBinding(), "excel/mex");

                host.Open();
                HostedInstance = host;
            }

            return HostedInstance;
        }

        public static void Stop()
        {
            if (HostedInstance != null)
            {
                HostedInstance.Close();
            }
        }
    }
}