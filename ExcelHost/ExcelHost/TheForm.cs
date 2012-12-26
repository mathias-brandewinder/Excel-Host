using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ServiceModel;
using System.ServiceModel.Channels;

namespace ExcelHost
{
    using ExcelService;

    public partial class TheForm : Form
    {
        private static IExcelService Service
        {
            get
            {
                var pipeFactory = new ChannelFactory<IExcelService>(new NetNamedPipeBinding(), new EndpointAddress("net.pipe://localhost/excel")); ;
                var proxy = pipeFactory.CreateChannel();
                return proxy;
            }
        }

        public TheForm()
        {
            InitializeComponent();
        }

        private void SendButton_Click(object sender, EventArgs e)
        {
            var data = new object[TestGrid.Rows.Count][];
            for (var row = 0; row < TestGrid.Rows.Count; ++row)
            {
                data[row] = new object[TestGrid.Columns.Count];
                for (var col = 0; col < TestGrid.Columns.Count; ++col)
                {
                    data[row][col] = TestGrid[col, row].Value;
                }
            }

            Service.DoMultipleStuff(data);
        }

        private void GetButton_Click(object sender, EventArgs e)
        {
            TestGrid.Rows.Clear();
            TestGrid.Refresh();

            var someData = Service.GrabMultipleIt();
            foreach (var row in someData)
            {
                TestGrid.Rows.Add(row);
            }
        }
    }
}
