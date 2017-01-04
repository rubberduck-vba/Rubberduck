using System;
using System.Windows.Forms;

namespace Rubberduck.UI.RegexAssistant
{
    public partial class RegexAssistantDialog : Form
    {
        public RegexAssistantDialog()
        {
            InitializeComponent();
            ViewModel = new RegexAssistantViewModel();
        }

        private RegexAssistantViewModel _viewModel;
        private RegexAssistantViewModel ViewModel { get { return _viewModel; }
        set
            {
                _viewModel = value;
                
                RegexAssistant.DataContext = _viewModel;
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
