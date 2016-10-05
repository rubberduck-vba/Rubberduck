using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Window : SafeComWrapper<Microsoft.Vbe.Interop.Window>, IEquatable<Window>
    {
        public Window(Microsoft.Vbe.Interop.Window window)
            : base(window)
        {
        }

        public int HWnd
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.HWnd); }
        }

        public IntPtr Handle()
        {
            return (IntPtr)HWnd;
        }

        public VBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : InvokeResult(() => ComObject.VBE)); }
        }

        public Windows Collection
        {
            get { return new Windows(IsWrappingNullReference ? null : InvokeResult(() => ComObject.Collection)); }
        }

        public string Caption
        {
            get { return IsWrappingNullReference ? string.Empty : InvokeResult(() => ComObject.Caption); }
        }

        public bool Visible
        {
            get { return !IsWrappingNullReference && InvokeResult(() => ComObject.Visible); }
            set { Invoke(() => ComObject.Visible = value); }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Left); }
            set { Invoke(() => ComObject.Left = value); }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Top); }
            set { Invoke(() => ComObject.Top = value); }
        }

        public int Width
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Width); }
            set { Invoke(() => ComObject.Width = value); }
        }

        public int Height
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => ComObject.Height); }
            set { Invoke(() => ComObject.Height = value); }
        }

        public WindowState WindowState
        {
            get { return IsWrappingNullReference ? 0 : InvokeResult(() => (WindowState)ComObject.WindowState); }
        }

        public WindowKind Type
        {
            get { return IsWrappingNullReference ? 0 : (WindowKind)InvokeResult(() => ComObject.Type); }
        }

        public LinkedWindows LinkedWindows
        {
            get { return new LinkedWindows(IsWrappingNullReference ? null : InvokeResult(() => ComObject.LinkedWindows)); }
        }

        public Window LinkedWindowFrame
        {
            get { return new Window(IsWrappingNullReference ? null : InvokeResult(() => ComObject.LinkedWindowFrame)); }
        }

        public void Close()
        {
            Invoke(() => ComObject.Close());
        }

        public void SetFocus()
        {
            Invoke(() => ComObject.SetFocus());
        }

        public void SetKind(WindowKind eKind)
        {
            Invoke(() => ComObject.SetKind((vbext_WindowType)eKind));
        }

        public void Detach()
        {
            Invoke(() => ComObject.Detach());
        }

        public void Attach(int lWindowHandle)
        {
            Invoke(() => ComObject.Attach(lWindowHandle));
        }
        
        public override void Release()
        {
            if (!IsWrappingNullReference)
            {
                LinkedWindowFrame.Release();
                Marshal.ReleaseComObject(ComObject);
            } 
        }

        public override bool Equals(SafeComWrapper<Microsoft.Vbe.Interop.Window> other)
        {
            return IsEqualIfNull(other) || (
                other != null 
                && (int)other.ComObject.Type == (int)Type 
                && other.ComObject.HWnd == HWnd);
        }

        public bool Equals(Window other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Window>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(HWnd, Type);
        }
    }
}