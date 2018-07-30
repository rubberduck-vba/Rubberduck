using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.WindowsApi
{
    //Stub for code pane replacement.  :-)
    internal class CodePaneSubclass : FocusSource, IWindowEventProvider
    {
        public event EventHandler<EventArgs> CaptionChanged;
        public event EventHandler<KeyPressEventArgs> KeyDown;
        private ICodePane _codePane;

        public ICodePane CodePane
        {
            get => _codePane;
            set
            {
                if (HasValidCodePane)
                {
                    _codePane.Dispose();          
                }

                _codePane = value;
            }
        }

        public bool HasValidCodePane
        {
            get
            {
                if (_codePane == null)
                {
                    return false;
                }

                try
                {
                    if (Marshal.GetIUnknownForObject(_codePane.Target) != IntPtr.Zero)
                    {
                        return true;
                    }

                    _codePane.Dispose();
                    _codePane = null;
                }
                catch
                {
                    // All paths leading to here mean that we need to ditch the held reference, and there
                    // isn't jack all that we can do about it.
                    _codePane = null;
                }
                return false;
            }
        }

        internal CodePaneSubclass(IntPtr hwnd, ICodePane pane) : base(hwnd)
        {
            _codePane = pane;
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
                    args = new KeyPressEventArgs(hWnd, wParam, lParam, (char)wParam);
                    OnKeyDown(args);
                    if (args.Handled) { return 0; }
                    break;
                case WM.KEYDOWN:
                    args = new KeyPressEventArgs(hWnd, wParam, lParam);
                    OnKeyDown(args);
                    if (args.Handled) { return 0; }
                    break;
                case WM.SETTEXT:
                    if (!HasValidCodePane)
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
                if (HasValidCodePane)
                {
                    _codePane.Dispose();
                    _codePane = null;
                }
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
