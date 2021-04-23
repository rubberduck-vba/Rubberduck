using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Rubberduck.Common;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;

namespace Rubberduck.UI.FindSymbol
{
    public class FindSymbolViewModel : INotifyPropertyChanged
    {
        private static readonly DeclarationType[] ExcludedTypes =
        {
            DeclarationType.Control,
            DeclarationType.Project
        };

        public FindSymbolViewModel(IEnumerable<Declaration> declarations, DeclarationIconCache cache)
        {
            _declarations = declarations.Where(declaration => !ExcludedTypes.Contains(declaration.DeclarationType)).ToList();
            _cache = cache;
            
            Search(string.Empty);
        }

        public event EventHandler<NavigateCodeEventArgs> Navigate;

        public bool CanExecute()
        {
            return _searchString?.Equals(_selectedItem?.IdentifierName, StringComparison.InvariantCultureIgnoreCase) ?? false;
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
            if (string.IsNullOrWhiteSpace(value))
            {
                MatchResults = new ObservableCollection<SearchResult>();
                return;
            }

            var results = GetSearchResultCollectionOfString(value);
            MatchResults = new ObservableCollection<SearchResult>(results);
        }

        private IEnumerable<SearchResult> GetSearchResultCollectionOfString(string value)
        {
            var lower = value.ToLowerInvariant();
            var results = _declarations
                .Where(declaration => string.IsNullOrEmpty(value) || declaration.IdentifierName.ToLowerInvariant().Contains(lower))
                .OrderBy(declaration => declaration.IdentifierName)
                .Take(80)
                .Select(declaration => new SearchResult(declaration, _cache[declaration]));

            return results;
        }

        private string _searchString;
        public string SearchString
        {
            get => _searchString;
            set
            {
                if (_searchString != value)
                {
                    _searchString = value;
                    Search(value);
                }
            }
        }

        private SearchResult _selectedItem;
        public SearchResult SelectedItem
        {
            get => _selectedItem;
            set 
            {
                if (_selectedItem != value)
                {
                    _selectedItem = value;
                    _searchString = value?.IdentifierName;
                    OnPropertyChanged();
                }
                if (_selectedItem != null)
                {
                    Execute();
                }
            }
        }

        private ObservableCollection<SearchResult> _matchResults;
        public ObservableCollection<SearchResult> MatchResults
        {
            get => _matchResults;
            set
            {
                _matchResults = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
