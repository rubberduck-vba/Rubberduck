using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Controls
{
    public class SearchResultsViewModel : ViewModelBase
    {
        private readonly string _header;

        public SearchResultsViewModel(string header, IEnumerable<SearchResultItem> searchResults)
        {
            _header = header;
            _searchResults = new ObservableCollection<SearchResultItem>(searchResults);
            _closeCommand = new DelegateCommand(ExecuteCloseCommand);
        }

        private readonly ObservableCollection<SearchResultItem> _searchResults;
        public ObservableCollection<SearchResultItem> SearchResults { get { return _searchResults; } }

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
    }
}