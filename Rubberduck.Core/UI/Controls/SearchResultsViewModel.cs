using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using NLog;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsViewModel : ViewModelBase, INavigateSelection
    {
        public SearchResultsViewModel(INavigateCommand navigateCommand, string header, Declaration target, IEnumerable<SearchResultItem> searchResults)
        {
            NavigateCommand = navigateCommand;
            Header = header;
            Target = target;
            SearchResultsSource = new CollectionViewSource();
            SearchResultsSource.GroupDescriptions.Add(new PropertyGroupDescription("ParentScope.QualifiedName.QualifiedModuleName.Name"));
            SearchResultsSource.SortDescriptions.Add(new SortDescription("ParentScope.QualifiedName.QualifiedModuleName.Name", ListSortDirection.Ascending));
            SearchResultsSource.SortDescriptions.Add(new SortDescription("Selection.StartLine", ListSortDirection.Ascending));
            SearchResultsSource.SortDescriptions.Add(new SortDescription("Selection.StartColumn", ListSortDirection.Ascending));

            SearchResults = new ObservableCollection<SearchResultItem>(searchResults);

            CloseCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), ExecuteCloseCommand);
        }

        private ObservableCollection<SearchResultItem> _searchResults;
        public ObservableCollection<SearchResultItem> SearchResults
        {
            get => _searchResults;
            set
            {
                _searchResults = value;

                SearchResultsSource.Source = _searchResults;

                OnPropertyChanged();
                OnPropertyChanged("SearchResultsSource");
            }
        }

        public CollectionViewSource SearchResultsSource { get; private set; }

        public string Header { get; }

        public CommandBase CloseCommand { get; }

        public Declaration Target { get; set; }

        private SearchResultItem _selectedItem;
        public SearchResultItem SelectedItem
        {
            get => _selectedItem;
            set
            {
                if (_selectedItem == value)
                {
                    return;
                }

                _selectedItem = value;
                OnPropertyChanged();
            }
        }

        private void ExecuteCloseCommand(object parameter)
        {
            OnClose();
        }

        public event EventHandler Close;
        private void OnClose()
        {
            Close?.Invoke(this, EventArgs.Empty);
        }

        public INavigateCommand NavigateCommand { get; }

        INavigateSource INavigateSelection.SelectedItem => SelectedItem;
    }
}
