using System.Windows.Forms;

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

        public string ClassId
        {
            get { return "BFD04A86-CACA-4F95-9656-A0BF7D3AE254"; }
        }

        public string Caption
        {
            get { return RubberduckUI.SearchResults_Caption; }
        }
    }
}
