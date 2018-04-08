using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using VB = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class Window : SafeComWrapper<VB.Window>, IWindow
    {
        public Window(VB.Window target, bool rewrapping = false)
            : base(target, rewrapping)
        {
        }

        public int HWnd => IsWrappingNullReference ? 0 : Target.HWnd;

        public IntPtr Handle()
        {
            return (IntPtr)HWnd;
        }

        public IVBE VBE => new VBE(IsWrappingNullReference ? null : Target.VBE);

        public IWindows Collection => new Windows(IsWrappingNullReference ? null : Target.Collection);

        public string Caption => IsWrappingNullReference ? string.Empty : Target.Caption;

        public bool IsVisible
        {
            get => !IsWrappingNullReference && Target.Visible;
            set => Target.Visible = value;
        }

        public int Left
        {
            get => IsWrappingNullReference ? 0 : Target.Left;
            set => Target.Left = value;
        }

        public int Top
        {
            get => IsWrappingNullReference ? 0 : Target.Top;
            set => Target.Top = value;
        }

        public int Width
        {
            get => IsWrappingNullReference ? 0 : Target.Width;
            set => Target.Width = value;
        }

        public int Height
        {
            get => IsWrappingNullReference ? 0 : Target.Height;
            set => Target.Height = value;
        }

        public WindowState WindowState => IsWrappingNullReference ? 0 : (WindowState)Target.WindowState;

        public WindowKind Type => IsWrappingNullReference ? 0 : (WindowKind)Target.Type;

        public ILinkedWindows LinkedWindows => new LinkedWindows(IsWrappingNullReference ? null : Target.LinkedWindows);

        public IWindow LinkedWindowFrame => new Window(IsWrappingNullReference ? null : Target.LinkedWindowFrame);

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