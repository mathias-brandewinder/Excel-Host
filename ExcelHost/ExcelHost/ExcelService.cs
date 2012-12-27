using System.Diagnostics;
using System.ServiceModel;
using ExcelService;

namespace ExcelHost
{
    using Microsoft.Office.Interop.Excel;

    internal class ExcelService : IExcelService
    {
        internal static Workbook Workbook { get; set; }
        private static ServiceHost HostedInstance { get; set; }
        private static readonly object Locker = new object();

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
            var fakeResult = new []
                                 {
                                     new object[] { "hi", "my", "name", "is", "hogmonaut" },
                                     new object[] { 1, 2, 3, 4, 5 },
                                     new object[] { "the", "answer", "is", 42, "!" },
                                     new object[] { "lolwhut", null, null, null, null }
                                 };

            return fakeResult;
        }

        public static void Start()
        {
            lock (Locker)
            {
                if (HostedInstance == null)
                {
                    HostedInstance = new ServiceHost(typeof(ExcelService));
                    HostedInstance.Open();
                }
            }
        }

        public static void Stop()
        {
            lock (Locker)
            {
                if (HostedInstance != null)
                {
                    HostedInstance.Abort();
                    HostedInstance = null;
                }
            }
        }
    }
}