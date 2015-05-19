using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.GoToAnything
{
    public partial class GoToAnythingDialog : Form
    {
        public GoToAnythingDialog(GoToAnythingViewModel viewModel)
            : this()
        {
            goToAnythingControl1.DataContext = viewModel;
        }

        public GoToAnythingDialog()
        {
            InitializeComponent();
        }
    }
}
