using System;
using System.Collections.Generic;

namespace Rubberduck.UI.Controls
{
    public interface ISearchResultsWindowViewModel
    {
        void AddTab(SearchResultsViewModel viewModel);
        event EventHandler LastTabClosed;
        IEnumerable<SearchResultsViewModel> Tabs { get; }
        SearchResultsViewModel SelectedTab { get; set; }
    }
}