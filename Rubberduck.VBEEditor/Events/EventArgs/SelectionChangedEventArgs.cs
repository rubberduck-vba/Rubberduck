using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class SelectionChangedEventArgs : EventArgs
    {
        public ICodePane CodePane { get; }

        public SelectionChangedEventArgs(ICodePane pane)
        {
            CodePane = pane;
        }
    }
}
