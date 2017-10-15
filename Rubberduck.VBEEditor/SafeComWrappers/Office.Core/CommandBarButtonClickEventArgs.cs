using System;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
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