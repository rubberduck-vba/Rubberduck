using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;

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