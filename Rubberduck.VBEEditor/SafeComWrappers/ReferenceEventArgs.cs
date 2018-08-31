using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class ReferenceEventArgs : EventArgs
    {
        public ReferenceEventArgs(IReference reference)
        {
            Reference = reference;
        }

        public IReference Reference { get; }
    }
}