using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Controls
{
    public class SearchResultItem : ViewModelBase, INavigateSource
    {
        private readonly NavigateCodeEventArgs _navigateArgs;
        private readonly Declaration _parentScopeDeclaration;
        private string _resultText;

        public SearchResultItem(Declaration parentScopeDeclaration, NavigateCodeEventArgs navigationInfo, string resultText)
        {
            _navigateArgs = navigationInfo;
            _parentScopeDeclaration = parentScopeDeclaration;
            _resultText = resultText;
        }

        public Declaration ParentScope { get { return _parentScopeDeclaration; } }
        public Selection Selection { get { return _navigateArgs.Selection; } }

        public string ResultText
        {
            get { return _resultText; }
            set
            {
                if (_resultText != value)
                {
                    _resultText = value;
                    OnPropertyChanged();
                }
            }
        }
        
        public NavigateCodeEventArgs GetNavigationArgs()
        {
            return _navigateArgs;
        }
    }
}
