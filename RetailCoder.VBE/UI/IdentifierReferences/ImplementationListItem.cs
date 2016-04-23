using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.IdentifierReferences
{
    public class ImplementationListItem
    {
        private readonly Declaration _declaration;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public ImplementationListItem(Declaration declaration, ICodePaneWrapperFactory wrapperFactory)
        {
            _declaration = declaration;
            _wrapperFactory = wrapperFactory;
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