using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Rubberduck.Annotations;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.GoToAnything
{
    public class GoToAnythingViewModel : INotifyPropertyChanged
    {
        public GoToAnythingViewModel(IEnumerable<Declaration> declarations)
        {
            _declarations = declarations;
            MatchResults = _declarations.OrderBy(declaration => declaration.IdentifierName)
                                        .Take(50)
                                        .Select(declaration => new SearchResult(declaration));
        }

        private readonly IEnumerable<Declaration> _declarations;

        private IEnumerable<SearchResult> _matchResults;

        public IEnumerable<SearchResult> MatchResults
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

    public class SearchResult
    {
        private readonly Declaration _declaration;

        public SearchResult(Declaration declaration)
        {
            _declaration = declaration;
        }

        public Declaration Declaration { get { return _declaration; } }
        public string IdentifierName { get { return _declaration.IdentifierName; } }
        public string Location { get { return _declaration.Scope; } }
    }
}
