using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class CommandBarButtonClickEventArgs : EventArgs
    {
        public CommandBarButtonClickEventArgs(ICommandBarButton control)
        {
            Control = control;
        }

        public ICommandBarButton Control { get; }

        public bool Cancel { get; set; }
        public bool Handled { get; set; } // Only used in VB6
    }
}