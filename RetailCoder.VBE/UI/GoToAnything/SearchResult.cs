using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.GoToAnything
{
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
