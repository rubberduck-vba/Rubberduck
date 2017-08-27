using System;

namespace AbstractRibbonDispatcher
{
    public class ChangedControlEventArgs : EventArgs {
        public ChangedControlEventArgs(string controlId) { ControlId = controlId; }
        public string ControlId { get; }
    }
}
