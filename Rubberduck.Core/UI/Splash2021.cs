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
    public partial class Splash2021 : Form
    {
        public Splash2021()
        {
            InitializeComponent();
#if DEBUG
            VersionLabel.Text = $"Version {GetType().Assembly.GetName().Version} (debug)";
#else
            VersionLabel.Text = $"Version {GetType().Assembly.GetName().Version}";
#endif
            VersionLabel.Parent = pictureBox1;
            VersionLabel.BackColor = Color.Transparent;
        }
    }
}
