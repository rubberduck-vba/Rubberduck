using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.FindSymbol
{
    public class SearchResult
    {
        private readonly Declaration _declaration;
        private readonly BitmapImage _icon;

        public SearchResult(Declaration declaration, BitmapImage icon)
        {
            _declaration = declaration;
            _icon = icon;
        }

        public Declaration Declaration { get { return _declaration; } }
        public string IdentifierName { get { return _declaration.IdentifierName; } }
        public string Location { get { return _declaration.Scope; } }

        public BitmapImage Icon { get { return _icon; } }

    }
}
