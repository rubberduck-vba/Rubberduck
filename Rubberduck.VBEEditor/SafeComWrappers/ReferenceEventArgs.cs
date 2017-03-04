using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class ReferenceEventArgs : EventArgs
    {
        private readonly IReference _reference;

        public ReferenceEventArgs(IReference reference)
        {
            _reference = reference;
        }

        public IReference Reference { get { return _reference; } }
    }
}