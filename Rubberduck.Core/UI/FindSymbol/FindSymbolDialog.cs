using System;
using System.Windows.Forms;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Resources;

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

        public FindSymbolDialog()
        {
            InitializeComponent();
            Text = $"Rubberduck - {RubberduckUI.FindSymbolDialog_Caption}";
        }


        public event EventHandler<NavigateCodeEventArgs> Navigate;
        private void viewModel_Navigate(object sender, NavigateCodeEventArgs e)
        {
            Navigate?.Invoke(this, e);
            Hide();
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