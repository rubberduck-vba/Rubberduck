using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using System.Windows.Input;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsViewModel : ViewModelBase, INavigateSelection
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly string _header;

        public SearchResultsViewModel(INavigateCommand navigateCommand, string header, IEnumerable<SearchResultItem> searchResults)
        {
            _navigateCommand = navigateCommand;
            _header = header;
            _searchResults = new ObservableCollection<SearchResultItem>(searchResults);
            _searchResultsSource = new CollectionViewSource();
            _searchResultsSource.Source = _searchResults;
            _searchResultsSource.GroupDescriptions.Add(new PropertyGroupDescription("QualifiedMemberName.QualifiedModuleName.Name"));
            _searchResultsSource.SortDescriptions.Add(new SortDescription("QualifiedMemberName.QualifiedModuleName.Name", ListSortDirection.Ascending));
            _closeCommand = new DelegateCommand(ExecuteCloseCommand);
        }

        private readonly ObservableCollection<SearchResultItem> _searchResults;
        public ObservableCollection<SearchResultItem> SearchResults { get { return _searchResults; } }

        private readonly CollectionViewSource _searchResultsSource;
        public CollectionViewSource SearchResultsSource { get { return _searchResultsSource; } }

        public string Header { get { return _header; } }

        private readonly ICommand _closeCommand;
        public ICommand CloseCommand { get { return _closeCommand; } }

        private SearchResultItem _selectedItem;
        public SearchResultItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                if (_selectedItem != value)
                {
                    _selectedItem = value;
                    OnPropertyChanged();
                }
            }
        }

        public INavigateCommand NavigateCommand { get {return _navigateCommand; } }

        private void ExecuteCloseCommand(object parameter)
        {
            OnClose();
        }

        public event EventHandler Close;
        private void OnClose()
        {
            var handler = Close;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        INavigateSource INavigateSelection.SelectedItem { get { return SelectedItem; } }
    }
}