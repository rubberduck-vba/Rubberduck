using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

namespace Rubberduck.API.VBA
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IIdentifierReferenceGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IIdentifierReference
    {
        [DispId(1)]
        Declaration Declaration { get; }
        [DispId(2)]
        Declaration ParentScope { get; }
        [DispId(3)]
        Declaration ParentNonScoping { get; }
        [DispId(4)]
        bool IsAssignment { get; }
        [DispId(5)]
        int StartLine { get; }
        [DispId(6)]
        int StartColumn { get; }
        [DispId(7)]
        int EndLine { get; }
        [DispId(8)]
        int EndColumn { get; }
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.IdentifierReferenceClassGuid),
        ProgId(RubberduckProgId.IdentifierReferenceProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IIdentifierReference)),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
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
