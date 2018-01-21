using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.IdentifierReferences
{
    public class IdentifierReferenceListItem
    {
        private readonly IdentifierReference _reference;

        public IdentifierReferenceListItem(IdentifierReference reference)
        {
            _reference = reference;
        }

        public IdentifierReference GetReferenceItem()
        {
            return _reference;
        }

        public QualifiedSelection Selection => new QualifiedSelection(_reference.QualifiedModuleName, _reference.Selection);
        public string IdentifierName => _reference.IdentifierName;

        public string DisplayString => $"{_reference.Context.Parent.GetText()} - {Selection.QualifiedName.ComponentName}, line {Selection.Selection.StartLine}";
    }
}
