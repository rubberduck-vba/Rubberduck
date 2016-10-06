using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class CommandBarButtonClickEventArgs : EventArgs
    {
        private readonly ICommandBarButton _control;

        internal CommandBarButtonClickEventArgs(ICommandBarButton control)
        {
            _control = control;
        }

        public ICommandBarButton Control { get { return _control; } }
        public bool Cancel { get; set; }
    }
}