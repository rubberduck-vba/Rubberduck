using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerMemberViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;

        public CodeExplorerMemberViewModel(Declaration declaration)
        {
            _declaration = declaration;
        }

        public string Name { get { return _declaration.IdentifierName; } }
        //public string Signature { get { return _declaration.IdentifierName; } }

    }
}