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

        private void viewModel_Close(object sender, EventArgs e)
        {
            RemoveTab(sender as SearchResultsViewModel);
        }

        public ObservableCollection<SearchResultsViewModel> Tabs => _tabs;

        private SearchResultsViewModel _selectedTab;
        public SearchResultsViewModel SelectedTab
        {
            get => _selectedTab;
            set
            {
                if (_selectedTab == value)
                {
                    return;
                }

                _selectedTab = value;
                OnPropertyChanged();
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
            LastTabClosed?.Invoke(this, EventArgs.Empty);
        }
    }
}
