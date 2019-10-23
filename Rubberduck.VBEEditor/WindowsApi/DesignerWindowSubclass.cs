using System;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.WindowsApi
{
    internal class DesignerWindowSubclass : FocusSource
    {
        //Stub for designer window replacement.  :-)
        internal DesignerWindowSubclass(IntPtr hwnd) : base(hwnd) { }

        public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            //Any time the selected control changes in the hosted userform, the F3 Overlay has to be redrawn.  This is a good proxy
            //for child control selections, so raise a focus event.
            if ((int) msg == (int)WM.ERASEBKGND)
            {
                DispatchFocusEvent(FocusType.GotFocus);
            }

            return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
        }
    }
}
