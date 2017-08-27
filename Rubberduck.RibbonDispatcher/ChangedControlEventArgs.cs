using System;

namespace Rubberduck.RibbonDispatcher
{
    public class ChangedControlEventArgs : EventArgs {
        public ChangedControlEventArgs(string controlId) { ControlId = controlId; }
        public string ControlId { get; }
    }
}
