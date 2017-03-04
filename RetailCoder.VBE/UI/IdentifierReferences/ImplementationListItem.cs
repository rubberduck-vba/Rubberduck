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

        public QualifiedSelection Selection { get { return new QualifiedSelection(_declaration.QualifiedName.QualifiedModuleName, _declaration.Selection); } }
        public string IdentifierName { get { return _declaration.IdentifierName; } }

        public string DisplayString
        {
            get
            {
                return string.Format("{0}, line {1}",
                    _declaration.Scope,
                    Selection.Selection.StartLine);
            }
        }
    }
}
