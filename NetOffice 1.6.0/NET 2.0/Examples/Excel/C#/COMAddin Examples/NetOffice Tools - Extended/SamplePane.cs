using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOfficeTools.ExtendedExcelCS2
{
    public partial class SamplePane : UserControl, NetOffice.ExcelApi.Tools.ITaskPane
    {
        public SamplePane()
        {
            InitializeComponent();
        }

        void NetOffice.ExcelApi.Tools.ITaskPane.OnConnection(NetOffice.ExcelApi.Application application, object[] customArguments)
        {

        }

        void NetOffice.ExcelApi.Tools.ITaskPane.OnDisconnection()
        {

        }
    }
}
