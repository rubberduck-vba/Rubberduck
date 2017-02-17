using System;

namespace Rubberduck.VBEditor.WindowsApi
{
    internal class DesignerWindowSubclass : FocusSource
    {
        //Stub for designer window replacement.  :-)
        internal DesignerWindowSubclass(IntPtr hwnd) : base(hwnd) { }
    }
}
