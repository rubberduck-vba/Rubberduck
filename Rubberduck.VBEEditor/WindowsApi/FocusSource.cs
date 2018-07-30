using System;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.WindowsApi
{
    internal abstract class FocusSource : SubclassingWindow, IFocusProvider 
    {
        protected FocusSource(IntPtr hwnd) : base(hwnd, hwnd) { }

        public event EventHandler<WindowChangedEventArgs> FocusChange;

        protected void OnFocusChange(WindowChangedEventArgs eventArgs)
        {
            FocusChange?.Invoke(this, eventArgs);
        }

        protected virtual void DispatchFocusEvent(FocusType type)
        {
            OnFocusChange(new WindowChangedEventArgs(Hwnd, type));
        }

        public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            switch ((uint)msg)
            {
                case (uint)WM.SETFOCUS:
                    DispatchFocusEvent(FocusType.GotFocus);
                    break;
                case (uint)WM.KILLFOCUS:
                    DispatchFocusEvent(FocusType.LostFocus);
                    break;
            }
            return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                FocusChange = delegate { };               
            }

            base.Dispose(disposing);
            _disposed = true;
        }
    }
}
