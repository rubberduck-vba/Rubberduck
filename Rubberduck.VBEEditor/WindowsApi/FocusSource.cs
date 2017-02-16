using System;
using Rubberduck.Common.WinAPI;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.WindowsApi
{
    internal abstract class FocusSource : SubclassingWindow, IWindowEventProvider
    {
        protected FocusSource(IntPtr hwnd) : base(hwnd, hwnd) { }

        public event EventHandler<WindowChangedEventArgs> FocusChange;
        private void OnFocusChange(WindowChangedEventArgs.FocusType type)
        {
            if (FocusChange != null)
            {
                var window = VBEEvents.GetWindowInfoFromHwnd(Hwnd);
                if (window == null)
                {
                    return;
                }
                FocusChange.Invoke(this, new WindowChangedEventArgs(window.Value.Hwnd, window.Value.Window, type));
            }
        } 

        public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            switch ((uint)msg)
            {
                case (uint)WM.SETFOCUS:
                    OnFocusChange(WindowChangedEventArgs.FocusType.GotFocus);
                    break;
                case (uint)WM.KILLFOCUS:
                    OnFocusChange(WindowChangedEventArgs.FocusType.LostFocus);
                    break;
            }
            return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
        }
    }
}
