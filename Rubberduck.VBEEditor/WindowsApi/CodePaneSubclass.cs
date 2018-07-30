using System;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.WindowsApi
{
    //Stub for code pane replacement.  :-)
    internal class CodePaneSubclass : FocusSource, IWindowEventProvider
    {       
        public event EventHandler CaptionChanged;
        public event EventHandler<KeyPressEventArgs> KeyDown;
        private ICodePane _codePane;

        /// <summary>
        /// The ICodePane associated with the message pump (if it has successfully been found).
        /// WARNING: Internal callers should NOT call *anything* on this object. Remember, you're in it's message pump here.
        /// External callers should NOT call .Dispose() on this object. That's the CodePaneSubclass's responsibility.
        /// </summary>
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

        /// <summary>
        /// Returns true if the Subclass is:
        /// 1.) Holding an ICodePane reference
        /// 2.) The held reference is pointed to a valid object (i.e. it has not been recycled). 
        /// </summary>
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
                    SubclassLogger.Warn($"{nameof(CodePaneSubclass)} failed to dispose of a held {nameof(ICodePane)} reference.");
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
