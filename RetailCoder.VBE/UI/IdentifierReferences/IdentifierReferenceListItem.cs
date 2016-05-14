using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferenceListItem
    {
        private readonly IdentifierReference _reference;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public IdentifierReferenceListItem(IdentifierReference reference, ICodePaneWrapperFactory wrapperFactory)
        {
            _reference = reference;
            _wrapperFactory = wrapperFactory;
        }

        public IdentifierReference GetReferenceItem()
        {
            return _reference;
        }

        public QualifiedSelection Selection { get { return new QualifiedSelection(_reference.QualifiedModuleName, _reference.Selection); } }
        public string IdentifierName { get { return _reference.IdentifierName; } }

        public string DisplayString 
        {
            get 
            { 
                return string.Format("{0} - {1}, line {2}", 
                    _reference.Context.Parent.GetText(), 
                    Selection.QualifiedName.ComponentName,
                    Selection.Selection.StartLine); 
            } 
        }
    }
}
