using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsWindowViewModel : ViewModelBase, ISearchResultsWindowViewModel
    {
        private readonly ObservableCollection<SearchResultsViewModel> _tabs = 
            new ObservableCollection<SearchResultsViewModel>();

        public void AddTab(SearchResultsViewModel viewModel)
        {
            viewModel.Close += viewModel_Close;
            _tabs.Add(viewModel);
        }

        void viewModel_Close(object sender, EventArgs e)
        {
            RemoveTab(sender as SearchResultsViewModel);
        }

        public ObservableCollection<SearchResultsViewModel> Tabs { get { return _tabs; } }

        private SearchResultsViewModel _selectedTab;
        public SearchResultsViewModel SelectedTab
        {
            get { return _selectedTab; }
            set
            {
                if (_selectedTab != value)
                {
                    _selectedTab = value;
                    OnPropertyChanged();
                }
            }
        }

        private void RemoveTab(SearchResultsViewModel viewModel)
        {
            if (viewModel != null)
            {
                _tabs.Remove(viewModel);
            }

            if (!_tabs.Any())
            {
                OnLastTabClosed();
            }
        }

        public event EventHandler LastTabClosed;
        private void OnLastTabClosed()
        {
            var handler = LastTabClosed;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }
    }
}
