using System;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class ReferenceEventArgs : EventArgs
    {
        public ReferenceEventArgs(ReferenceInfo reference, ReferenceKind type)
        {
            Reference = reference;
            Type = type;
        }

        public ReferenceInfo Reference { get; }
        public ReferenceKind Type { get; }
    }
}