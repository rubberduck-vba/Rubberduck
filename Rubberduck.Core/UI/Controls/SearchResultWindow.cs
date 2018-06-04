using System;
using System.Windows.Forms;
using Rubberduck.Resources;

namespace Rubberduck.UI.Controls
{
    public partial class SearchResultWindow : UserControl, IDockableUserControl
    {
        public SearchResultWindow()
        {
            InitializeComponent();
        }

        private ISearchResultsWindowViewModel _viewModel;
        public ISearchResultsWindowViewModel ViewModel
        {
            get { return _viewModel; }
            set
            {
                _viewModel = value;
                searchView1.DataContext = _viewModel;
            }
        }

        private readonly string RandomGuid = Guid.NewGuid().ToString();
        string IDockableUserControl.GuidIdentifier => RandomGuid;

        public string Caption
        {
            get { return RubberduckUI.SearchResults_Caption; }
        }
    }
}
