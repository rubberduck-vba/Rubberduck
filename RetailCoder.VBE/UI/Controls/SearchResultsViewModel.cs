using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsViewModel : ViewModelBase, INavigateSelection
    {
        private readonly INavigateCommand _navigateCommand;
        private readonly string _header;

        public SearchResultsViewModel(INavigateCommand navigateCommand, string header, Declaration target, IEnumerable<SearchResultItem> searchResults)
        {
            _navigateCommand = navigateCommand;
            _header = header;
            Target = target;
            SearchResultsSource = new CollectionViewSource();
            SearchResultsSource.GroupDescriptions.Add(new PropertyGroupDescription("ParentScope.QualifiedName.QualifiedModuleName.Name"));
            SearchResultsSource.SortDescriptions.Add(new SortDescription("ParentScope.QualifiedName.QualifiedModuleName.Name", ListSortDirection.Ascending));
            SearchResultsSource.SortDescriptions.Add(new SortDescription("Selection.StartLine", ListSortDirection.Ascending));
            SearchResultsSource.SortDescriptions.Add(new SortDescription("Selection.StartColumn", ListSortDirection.Ascending));

            SearchResults = new ObservableCollection<SearchResultItem>(searchResults);

            _closeCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCloseCommand);
        }

        private ObservableCollection<SearchResultItem> _searchResults;
        public ObservableCollection<SearchResultItem> SearchResults
        {
            get { return _searchResults; }
            set
            {
                _searchResults = value;

                SearchResultsSource.Source = _searchResults;

                OnPropertyChanged();
                OnPropertyChanged("SearchResultsSource");
            }
        }

        public CollectionViewSource SearchResultsSource { get; private set; }

        public string Header { get { return _header; } }

        private readonly CommandBase _closeCommand;
        public CommandBase CloseCommand { get { return _closeCommand; } }

        public Declaration Target { get; set; }

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

        public INavigateCommand NavigateCommand { get { return _navigateCommand; } }
        INavigateSource INavigateSelection.SelectedItem { get { return SelectedItem; } }
    }
}
