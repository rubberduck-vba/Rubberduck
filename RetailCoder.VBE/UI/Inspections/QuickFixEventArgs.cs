using System;

namespace Rubberduck.UI.Inspections
{
    public class QuickFixEventArgs : EventArgs
    {
        public QuickFixEventArgs(Action quickFix)
        {
            QuickFix = quickFix;
        }

        public Action QuickFix { get; }
    }
}
