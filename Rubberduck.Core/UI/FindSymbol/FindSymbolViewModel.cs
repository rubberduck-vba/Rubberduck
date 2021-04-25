using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using NLog;
using Rubberduck.Common;
using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.FindSymbol
{
    public class FindSymbolViewModel : INotifyPropertyChanged
    {
        private static readonly DeclarationType[] ExcludedTypes =
        {
            DeclarationType.Control,
            DeclarationType.Project
        };

        private static readonly ILogger Logger = LogManager.GetCurrentClassLogger();

        public FindSymbolViewModel(IEnumerable<Declaration> declarations)
        {
            _declarations = declarations.Where(declaration => !ExcludedTypes.Contains(declaration.DeclarationType)).ToList();
            GoCommand = new DelegateCommand(Logger, ExecuteGoCommand, CanExecuteGoCommand);
            Search(string.Empty);
        }

        public event EventHandler<NavigateCodeEventArgs> Navigate;

        public ICommand GoCommand { get; }

        private bool CanExecuteGoCommand(object param) => _searchString?.Equals(_selectedItem?.IdentifierName, StringComparison.InvariantCultureIgnoreCase) ?? false;

        private void ExecuteGoCommand(object param) => OnNavigate();


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
                .Take(30)
                .Select(declaration => new SearchResult(declaration));

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
                    OnPropertyChanged();

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
