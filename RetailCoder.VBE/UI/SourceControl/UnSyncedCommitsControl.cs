﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.SourceControl
{
    public partial class UnSyncedCommitsControl : UserControl, IUnSyncedCommitsView
    {
        public UnSyncedCommitsControl()
        {
            InitializeComponent();
        }
    }
}
