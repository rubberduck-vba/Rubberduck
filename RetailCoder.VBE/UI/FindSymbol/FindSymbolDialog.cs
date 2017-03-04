using System;
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

            Text = string.Format("Rubberduck - {0}", RubberduckUI.FindSymbolDialog_Caption);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}
