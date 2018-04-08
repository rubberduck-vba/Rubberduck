using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.FindSymbol
{
    public class SearchResult
    {
        public SearchResult(Declaration declaration, BitmapImage icon)
        {
            Declaration = declaration;
            Icon = icon;
        }

        public Declaration Declaration { get; }

        public string IdentifierName => Declaration.IdentifierName;
        public string Location => Declaration.Scope;

        public BitmapImage Icon { get; }
    }
}
