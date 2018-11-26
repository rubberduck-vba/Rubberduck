using System;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.WindowsApi
{
    //Stub for code pane replacement.  :-)
    internal class CodePaneSubclass : VbeAttachableSubclass<ICodePane>, IWindowEventProvider
    {       
        public event EventHandler CaptionChanged;
        public event EventHandler<KeyPressEventArgs> KeyDown;
 
        internal CodePaneSubclass(IntPtr hwnd, ICodePane pane) : base(hwnd)
        {
            VbeObject = pane;
        }

        protected void OnKeyDown(KeyPressEventArgs eventArgs)
        {
            KeyDown?.Invoke(this, eventArgs);
        }

        public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            KeyPressEventArgs args;
            switch ((WM)msg)
            {
                case WM.CHAR:
                    args = new KeyPressEventArgs(hWnd, wParam, lParam);
                    if (args.Character != '\r' && args.Character != '\n' && args.Character != '\b')
                    {
                        OnKeyDown(args);
                        if (args.Handled) { return 0; }
                    }
                    break;
                case WM.KEYDOWN:
                    args = new KeyPressEventArgs(hWnd, wParam, lParam, true);
                    // The only keydown we care about that doesn't generate a WM_CHAR is Delete, and the VBE handles Enter & backspace in WM_KEYDOWN, 
                    // so we need to handle them first (otherwise it will already be in the code pane when the managed event is handled).
                    if (args.IsDelete || args.Character == '\r' || args.Character == '\b')
                    {
                        OnKeyDown(args);
                        if (args.Handled) { return 0; }
                    }
                    break;
                case WM.SETTEXT:
                    if (!HasValidVbeObject)
                    {
                        CaptionChanged?.Invoke(this, null);
                    }
                    break;
            }
            return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                CaptionChanged = delegate { };
                KeyDown = delegate { };
            }

            base.Dispose(disposing);
            _disposed = true;
        }

        protected override void DispatchFocusEvent(FocusType type)
        {
            OnFocusChange(new WindowChangedEventArgs(Hwnd, type));
        }
    }
}
