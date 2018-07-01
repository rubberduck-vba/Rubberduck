using System;
using System.Windows.Forms;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.WindowsApi
{
    public class IntelliSenseEventArgs : EventArgs
    {
        public static IntelliSenseEventArgs Shown => new IntelliSenseEventArgs(true);
        public static IntelliSenseEventArgs Hidden => new IntelliSenseEventArgs(false);
        internal IntelliSenseEventArgs(bool visible)
        {
            Visible = visible;
        }

        public bool Visible { get; }
    }

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

        public bool Handled { get; set; }

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
            KeyPressEventArgs args;
            switch ((WM)msg)
            {
                case WM.CHAR:
                    args = new KeyPressEventArgs(hWnd, wParam, lParam, (char)wParam);
                    OnKeyDown(args);
                    if (args.Handled) { return 0; }
                    break;
                case WM.KEYDOWN:
                    args = new KeyPressEventArgs(hWnd, wParam, lParam);
                    OnKeyDown(args);
                    if (args.Handled) { return 0; }
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
