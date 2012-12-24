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
        public TheForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var message = this.textBox1.Text;
            var pipeFactory = new ChannelFactory<IExcelService>(new NetNamedPipeBinding(), new EndpointAddress("net.pipe://localhost/excel")); ;
            var proxy = pipeFactory.CreateChannel();
            proxy.DoStuff(message);
        }
    }
}
