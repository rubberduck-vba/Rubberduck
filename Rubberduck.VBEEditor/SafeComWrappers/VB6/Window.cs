using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class Window : SafeComWrapper<VB.Window>, IWindow
    {
        public Window(VB.Window window)
            : base(window)
        {
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
            Target.SetKind((VB.vbext_WindowType)eKind);
        }

        public void Detach()
        {
            Target.Detach();
        }

        public void Attach(int lWindowHandle)
        {
            Target.Attach(lWindowHandle);
        }
        
        //public override void Release(bool final = false)
        //{
        //    if (!IsWrappingNullReference)
        //    {
        //        LinkedWindowFrame.Release();
        //        base.Release(final);
        //    } 
        //}

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