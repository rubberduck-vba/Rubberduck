using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rubberduck.UI.FindSymbol
{
    public partial class FindSymbolDialog : Form
    {
        public FindSymbolDialog(FindSymbolViewModel viewModel)
            : this()
        {
            findSymbolControl1.DataContext = viewModel;
            viewModel.Navigate += viewModel_Navigate;
        }

        public event EventHandler<NavigateCodeEventArgs> Navigate;
        private void viewModel_Navigate(object sender, NavigateCodeEventArgs e)
        {
            var handler = Navigate;
            if (handler != null)
            {
                handler(this, e);
                Hide();
            }
        }

        public FindSymbolDialog()
        {
            InitializeComponent();
        }
    }
}
