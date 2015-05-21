using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Rubberduck.Annotations;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.FindSymbol
{
    public class FindSymbolViewModel : INotifyPropertyChanged
    {
        public FindSymbolViewModel(IEnumerable<Declaration> declarations)
        {
            _declarations = declarations;
            var initialResults = _declarations.OrderBy(declaration => declaration.IdentifierName.ToLowerInvariant())
                .Select(declaration => new SearchResult(declaration));

            MatchResults = new ObservableCollection<SearchResult>(initialResults);
        }

        public event EventHandler<NavigateCodeEventArgs> Navigate;

        public bool CanExecute()
        {
            return _selectedItem != null;
        }

        public void Execute()
        {
            OnNavigate();
        }

        public void OnNavigate()
        {
            var handler = Navigate;
            if (handler != null && _selectedItem != null)
            {
                var arg = new NavigateCodeEventArgs(_selectedItem.Declaration);
                handler(this, arg);
            }
        }

        private readonly IEnumerable<Declaration> _declarations;

        private void Search(string value)
        {
            var lower = value.ToLowerInvariant();
            var results = _declarations.Where(
                declaration => declaration.IdentifierName.ToLowerInvariant().Contains(lower))
                .OrderBy(declaration => declaration.IdentifierName.ToLowerInvariant())
                .Select(declaration => new SearchResult(declaration));

            MatchResults = new ObservableCollection<SearchResult>(results);
        }

        private string _searchString;

        public string SearchString
        {
            get { return _searchString; }
            set
            {
                _searchString = value; 
                Search(value);
            }
        }

        private SearchResult _selectedItem;

        public SearchResult SelectedItem
        {
            get { return _selectedItem; }
            set 
            { 
                _selectedItem = value; 
                OnPropertyChanged();
            }
        }

        private ObservableCollection<SearchResult> _matchResults;

        public ObservableCollection<SearchResult> MatchResults
        {
            get { return _matchResults; }
            set { _matchResults = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}