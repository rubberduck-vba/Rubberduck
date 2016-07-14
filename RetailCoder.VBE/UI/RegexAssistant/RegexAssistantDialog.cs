using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    }
}
