using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Native;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Window : SafeComWrapper<VB.Window>, IWindow
    {
        public Window(VB.Window window)
            : base(window)
        {
            _events = WinEventProc;
        }

        //ReSharper disable once PrivateFieldCanBeConvertedToLocalVariable
        private readonly WinEvents.WinEventDelegate _events;
        private IntPtr _hook = IntPtr.Zero;
        private IntPtr _hwnd = IntPtr.Zero;

        public IntPtr RealHwnd
        {
            get { return _hwnd; }
            set
            {
                if (value == _hwnd)
                {
                    return;
                }
                //Are we hooked to a different hwnd?
                if (_hook != IntPtr.Zero)
                {
                    WinEvents.UnhookWinEvent(_hook);
                }
                _hwnd = value;
                //May as well set the WinEventHook too now that we have an hwnd.
                if (_hwnd == IntPtr.Zero)
                {
                    return;
                }
                //Just grab everything for now.  We can narrow this once we determine what we care about.
                _hook = WinEvents.SetWinEventHook((uint)WinEvents.EventConstant.EVENT_MIN,
                    (uint)WinEvents.EventConstant.EVENT_MAX, IntPtr.Zero, Marshal.GetFunctionPointerForDelegate(_events), 0, 0,
                    (uint)WinEvents.WinEventFlags.WINEVENT_OUTOFCONTEXT);
            }
        }

        public int HWnd
        {
            get { return IsWrappingNullReference ? 0 : Target.HWnd; }
        }

        public IntPtr Handle()
        {
            return (IntPtr)HWnd;
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : Target.VBE); }
        }
        
        public IWindows Collection
        {
            get { return new Windows(IsWrappingNullReference ? null : Target.Collection); }
        }

        public string Caption
        {
            get { return IsWrappingNullReference ? string.Empty : Target.Caption; }
        }

        public bool IsVisible
        {
            get { return !IsWrappingNullReference && Target.Visible; }
            set { Target.Visible = value; }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : Target.Left; }
            set { Target.Left = value; }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : Target.Top; }
            set { Target.Top = value; }
        }

        public int Width
        {
            get { return IsWrappingNullReference ? 0 : Target.Width; }
            set { Target.Width = value; }
        }

        public int Height
        {
            get { return IsWrappingNullReference ? 0 : Target.Height; }
            set { Target.Height = value; }
        }

        public WindowState WindowState
        {
            get { return IsWrappingNullReference ? 0 : (WindowState)Target.WindowState; }
        }

        public WindowKind Type
        {
            get { return IsWrappingNullReference ? 0 : (WindowKind)Target.Type; }
        }

        public ILinkedWindows LinkedWindows
        {
            get { return new LinkedWindows(IsWrappingNullReference ? null : Target.LinkedWindows); }
        }

        public IWindow LinkedWindowFrame
        {
            get { return new Window(IsWrappingNullReference ? null : Target.LinkedWindowFrame); }
        }

        public void Close()
        {
            Target.Close();
        }

        public void SetFocus()
        {
            Target.SetFocus();
        }

        public void SetKind(WindowKind eKind)
        {
            Target.SetKind((vbext_WindowType)eKind);
        }

        public void Detach()
        {
            Target.Detach();
        }

        public void Attach(int lWindowHandle)
        {
            Target.Attach(lWindowHandle);
        }
        
        public override void Release(bool final = false)
        {
            if (_hook != IntPtr.Zero)
            {
                WinEvents.UnhookWinEvent(_hook);
                _hook = IntPtr.Zero;
            }
            if (!IsWrappingNullReference)
            {
                LinkedWindowFrame.Release();
                base.Release(final);
            } 
        }

        public event EventHandler<EventArgs> Activate;
        public event EventHandler<EventArgs> Deactivate;

        private bool _hasfocus;
        protected void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, uint idObject, uint idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if ((WinEvents.EventConstant)eventType == WinEvents.EventConstant.EVENT_SYSTEM_FOREGROUND)
            {
                _hasfocus = hwnd == _hwnd;
                OnFocusChange();
            }
        }

        protected virtual void OnFocusChange()
        {
            var handler = _hasfocus ? Activate : Deactivate;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        public override bool Equals(ISafeComWrapper<VB.Window> other)
        {
            return IsEqualIfNull(other) || (
                other != null 
                && (int)other.Target.Type == (int)Type 
                && other.Target.HWnd == HWnd);
        }

        public bool Equals(IWindow other)
        {
            return Equals(other as SafeComWrapper<VB.Window>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : HashCode.Compute(HWnd, Type);
        }
    }
}