using System;
using System.Windows.Forms;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.WindowsApi
{
    public class KeyPressEventArgs
    {
        public KeyPressEventArgs(IntPtr hwnd, IntPtr wParam, IntPtr lParam, char character = default)
        {
            Hwnd = hwnd;
            WParam = wParam;
            LParam = lParam;
            Character = character;
            if (character == default(char))
            {
                Key = (Keys)wParam;
            }
            else
            {
                IsCharacter = true;
            }
        }

        public bool IsCharacter { get; }
        public IntPtr Hwnd { get; }
        public IntPtr WParam { get; }
        public IntPtr LParam { get; }

        public char Character { get; }
        public Keys Key { get; }
    }

    //Stub for code pane replacement.  :-)
    internal class CodePaneSubclass : FocusSource
    {
        public ICodePane CodePane { get; }

        internal CodePaneSubclass(IntPtr hwnd, ICodePane pane) : base(hwnd)
        {
            CodePane = pane;
        }

        public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            switch ((WM)msg)
            {
                case WM.CHAR:
                    OnKeyDown(new KeyPressEventArgs(hWnd, wParam, lParam, (char)wParam));
                    break;
                case WM.KEYDOWN:
                    OnKeyDown(new KeyPressEventArgs(hWnd, wParam, lParam));
                    break;
            }
            return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
        }

        protected override void DispatchFocusEvent(FocusType type)
        {
            var window = VBENativeServices.GetWindowInfoFromHwnd(Hwnd);
            if (!window.HasValue)
            {
                return;
            }
            OnFocusChange(new WindowChangedEventArgs(window.Value.Hwnd, window.Value.Window, CodePane, type));
        }
    }
}
