using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;

namespace Rubberduck.UI.FindSymbol
{
    public class FindSymbolViewModel : INotifyPropertyChanged
    {
        private static readonly DeclarationType[] ExcludedTypes =
        {
            DeclarationType.Control, 
            DeclarationType.ModuleOption,
            DeclarationType.Project
        };

        public FindSymbolViewModel(IEnumerable<Declaration> declarations, DeclarationIconCache cache)
        {
            _declarations = declarations;
            _cache = cache;
            var initialResults = _declarations
                .Where(declaration => !ExcludedTypes.Contains(declaration.DeclarationType))
                .OrderBy(declaration => declaration.IdentifierName.ToLowerInvariant())
                .Select(declaration => new SearchResult(declaration, cache[declaration]))
                .ToList();

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
        private readonly DeclarationIconCache _cache;

        private void Search(string value)
        {
            var lower = value.ToLowerInvariant();
            var results = _declarations
                .Where(declaration => !ExcludedTypes.Contains(declaration.DeclarationType)
                                        && (string.IsNullOrEmpty(value) || declaration.IdentifierName.ToLowerInvariant().Contains(lower)))
                .OrderBy(declaration => declaration.IdentifierName.ToLowerInvariant())
                .Select(declaration => new SearchResult(declaration, _cache[declaration]))
                .ToList();

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
