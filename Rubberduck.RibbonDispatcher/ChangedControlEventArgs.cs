using System;

namespace Rubberduck.RibbonDispatcher
{
    [CLSCompliant(true)]
    public class ChangedControlEventArgs : EventArgs {
        public ChangedControlEventArgs(string controlId) { ControlId = controlId; }
        public string ControlId { get; }
    }
}
