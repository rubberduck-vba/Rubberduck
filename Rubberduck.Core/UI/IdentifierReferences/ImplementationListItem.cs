using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ImplementationListItem
    {
        private readonly Declaration _declaration;

        public ImplementationListItem(Declaration declaration)
        {
            _declaration = declaration;
        }

        public Declaration GetDeclaration()
        {
            return _declaration;
        }

        public QualifiedSelection Selection => new QualifiedSelection(_declaration.QualifiedName.QualifiedModuleName, _declaration.Selection);
        public string IdentifierName => _declaration.IdentifierName;

        public string DisplayString => $"{_declaration.Scope}, line {Selection.Selection.StartLine}";
    }
}
