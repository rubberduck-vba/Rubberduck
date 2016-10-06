using System;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class ReferenceEventArgs : EventArgs
    {
        private readonly Reference _reference;

        public ReferenceEventArgs(Reference reference)
        {
            _reference = reference;
        }

        public Reference Reference { get { return _reference; } }
    }
}