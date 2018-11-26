using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsWindowViewModel : ViewModelBase, ISearchResultsWindowViewModel
    {
        public void AddTab(SearchResultsViewModel viewModel)
        {
            viewModel.Close += viewModel_Close;
            Tabs.Add(viewModel);
        }

        private void viewModel_Close(object sender, EventArgs e)
        {
            RemoveTab(sender as SearchResultsViewModel);
        }

        public ObservableCollection<SearchResultsViewModel> Tabs { get; } = new ObservableCollection<SearchResultsViewModel>();

        private SearchResultsViewModel _selectedTab;
        public SearchResultsViewModel SelectedTab
        {
            get => _selectedTab;
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
                Tabs.Remove(viewModel);
            }

            if (!Tabs.Any())
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
