﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOfficeTools.ExtendedWordCS2
{
    public partial class SamplePane : UserControl, NetOffice.WordApi.Tools.ITaskPane
    {
        public SamplePane()
        {
            InitializeComponent();
        }

        void NetOffice.WordApi.Tools.ITaskPane.OnConnection(NetOffice.WordApi.Application application, object[] customArguments)
        {

        }

        void NetOffice.WordApi.Tools.ITaskPane.OnDisconnection()
        {

        }
    }
}
