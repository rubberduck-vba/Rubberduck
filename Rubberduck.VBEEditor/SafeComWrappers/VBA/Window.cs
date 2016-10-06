using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class Window : SafeComWrapper<Microsoft.Vbe.Interop.Window>, IWindow
    {
        public Window(Microsoft.Vbe.Interop.Window window)
            : base(window)
        {
        }

        public int HWnd
        {
            get { return IsWrappingNullReference ? 0 : ComObject.HWnd; }
        }

        public IntPtr Handle()
        {
            return (IntPtr)HWnd;
        }

        public IVBE VBE
        {
            get { return new VBE(IsWrappingNullReference ? null : ComObject.VBE); }
        }
        
        public IWindows Collection
        {
            get { return new Windows(IsWrappingNullReference ? null : ComObject.Collection); }
        }

        public string Caption
        {
            get { return IsWrappingNullReference ? string.Empty : ComObject.Caption; }
        }

        public bool IsVisible
        {
            get { return !IsWrappingNullReference && ComObject.Visible; }
            set { ComObject.Visible = value; }
        }

        public int Left
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Left; }
            set { ComObject.Left = value; }
        }

        public int Top
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Top; }
            set { ComObject.Top = value; }
        }

        public int Width
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Width; }
            set { ComObject.Width = value; }
        }

        public int Height
        {
            get { return IsWrappingNullReference ? 0 : ComObject.Height; }
            set { ComObject.Height = value; }
        }

        public WindowState WindowState
        {
            get { return IsWrappingNullReference ? 0 : (WindowState)ComObject.WindowState; }
        }

        public WindowKind Type
        {
            get { return IsWrappingNullReference ? 0 : (WindowKind)ComObject.Type; }
        }

        public ILinkedWindows LinkedWindows
        {
            get { return new LinkedWindows(IsWrappingNullReference ? null : ComObject.LinkedWindows); }
        }

        public IWindow LinkedWindowFrame
        {
            get { return new Window(IsWrappingNullReference ? null : ComObject.LinkedWindowFrame); }
        }

        public void Close()
        {
            ComObject.Close();
        }

        public void SetFocus()
        {
            ComObject.SetFocus();
        }

        public void SetKind(WindowKind eKind)
        {
            ComObject.SetKind((vbext_WindowType)eKind);
        }

        public void Detach()
        {
            ComObject.Detach();
        }

        public void Attach(int lWindowHandle)
        {
            ComObject.Attach(lWindowHandle);
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

        public bool Equals(IWindow other)
        {
            return Equals(other as SafeComWrapper<Microsoft.Vbe.Interop.Window>);
        }

        public override int GetHashCode()
        {
            return IsWrappingNullReference ? 0 : ComputeHashCode(HWnd, Type);
        }
    }
}