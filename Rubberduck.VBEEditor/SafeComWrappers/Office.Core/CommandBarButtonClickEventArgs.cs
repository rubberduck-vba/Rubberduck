using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core
{
    public class CommandBarButtonClickEventArgs : EventArgs
    {
        private readonly CommandBarButton _control;

        internal CommandBarButtonClickEventArgs(CommandBarButton control)
        {
            _control = control;
        }

        public CommandBarButton Control { get { return _control; } }
        public bool Cancel { get; set; }
    }
}