using System;

namespace Rubberduck.UI.Inspections
{
    public class QuickFixEventArgs : EventArgs
    {
        private readonly Action _quickFix;

        public QuickFixEventArgs(Action quickFix)
        {
            _quickFix = quickFix;
        }

        public Action QuickFix
        {
            get { return _quickFix; }
        }
    }
}
