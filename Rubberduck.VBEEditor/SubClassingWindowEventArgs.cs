using System;
using System.Windows.Forms;

namespace Rubberduck.VBEditor
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