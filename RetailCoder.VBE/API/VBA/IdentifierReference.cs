using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.API.VBA
{
    [ComVisible(true)]
    public interface IIdentifierReference
    {
        Declaration Declaration { get; }
        Declaration ParentScope { get; }
        Declaration ParentNonScoping { get; }
        bool IsAssignment { get; }
        int StartLine { get; }
        int StartColumn { get; }
        int EndLine { get; }
        int EndColumn { get; }
    }

    [ComVisible(true)]
    [Guid(RubberduckGuid.IdentifierReferenceClassGuid)]
    [ProgId(RubberduckProgId.IdentifierReferenceProgId)]
    [ComDefaultInterface(typeof(IIdentifierReference))]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public class IdentifierReference : IIdentifierReference
    {
        private readonly Parsing.Symbols.IdentifierReference _reference;

        public IdentifierReference(Parsing.Symbols.IdentifierReference reference)
        {
            _reference = reference;
        }

        private Declaration _declaration;
        public Declaration Declaration => _declaration ?? (_declaration = new Declaration(_reference.Declaration));

        private Declaration _parentScoping;
        public Declaration ParentScope => _parentScoping ?? (_parentScoping = new Declaration(_reference.ParentScoping));

        private Declaration _parentNonScoping;
        public Declaration ParentNonScoping => _parentNonScoping ?? (_parentNonScoping = new Declaration(_reference.ParentNonScoping));

        public bool IsAssignment => _reference.IsAssignment;

        public int StartLine => _reference.Selection.StartLine;
        public int EndLine => _reference.Selection.EndLine;
        public int StartColumn => _reference.Selection.StartColumn;
        public int EndColumn => _reference.Selection.EndColumn;
    }
}
