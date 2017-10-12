using System;
using Rubberduck.VBEditor.SafeComWrappers.Office.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office
{
    public class CommandBarButtonClickEventArgs : EventArgs
    {
        internal CommandBarButtonClickEventArgs(ICommandBarButton control)
        {
            Control = control;
        }

        public ICommandBarButton Control { get; }

        public bool Cancel { get; set; }
    }
}