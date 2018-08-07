using System;

namespace Rubberduck.VBEditor.Events
{
    public class SubClassingWindowEventArgs : EventArgs
    {
        public IntPtr LParam { get; }

        public bool Closing { get; set; }

        public SubClassingWindowEventArgs(IntPtr lparam)
        {
            LParam = lparam;
        }
    }
}