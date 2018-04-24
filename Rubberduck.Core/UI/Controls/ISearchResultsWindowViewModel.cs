using System;
using System.Collections.ObjectModel;

namespace Rubberduck.UI.Controls
{
    public interface ISearchResultsWindowViewModel
    {
        void AddTab(SearchResultsViewModel viewModel);
        event EventHandler LastTabClosed;
        ObservableCollection<SearchResultsViewModel> Tabs { get; }
        SearchResultsViewModel SelectedTab { get; set; }
    }
}
