using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    public partial class Splash : Form
    {
        public Splash()
        {
            InitializeComponent();
        }

        public string Version
        {
            get { return VersionLabel.Text; }
            set { VersionLabel.Text = value; }
        }
    }
}
