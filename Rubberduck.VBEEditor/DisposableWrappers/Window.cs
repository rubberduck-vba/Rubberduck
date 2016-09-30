using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class Window : WrapperBase<Microsoft.Vbe.Interop.Window>, IDisposable
    {
        public Window(Microsoft.Vbe.Interop.Window window)
            : base(window)
        {
        }

        public void Close()
        {
            ThrowIfDisposed();
            InvokeMember(() => Item.Close());
        }

        public void SetFocus()
        {
            ThrowIfDisposed();
            InvokeMember(() => Item.SetFocus());
        }

        public void SetKind(WindowKind eKind)
        {
            ThrowIfDisposed();
            InvokeMember(kind => Item.SetKind((vbext_WindowType)kind), eKind);
        }

        public void Detach()
        {
            ThrowIfDisposed();
            InvokeMember(() => Item.Detach());
        }

        public void Attach(int lWindowHandle)
        {
            ThrowIfDisposed();
            InvokeMember(handle => Item.Attach(handle), lWindowHandle);
        }

        public VBE VBE
        {
            get
            {
                ThrowIfDisposed(); 
                return new VBE(InvokeMemberValue(() => Item.VBE));
            }
        }

        public Windows Collection
        {
            get
            {
                ThrowIfDisposed();
                return new Windows(InvokeMemberValue(() => Item.Collection));
            }
        }

        public string Caption
        {
            get
            {
                ThrowIfDisposed(); 
                return InvokeMemberValue(() => Item.Caption);
            }
        }

        public bool Visible
        {
            get
            {
                ThrowIfDisposed(); 
                return InvokeMemberValue(() => Item.Visible);
            }
        }

        public int Left
        {
            get
            {
                ThrowIfDisposed(); 
                return InvokeMemberValue(() => Item.Left);
            }
        }

        public int Top
        {
            get
            {
                ThrowIfDisposed(); 
                return InvokeMemberValue(() => Item.Top);
            }
        }

        public int Width
        {
            get
            {
                ThrowIfDisposed(); 
                return InvokeMemberValue(() => Item.Width);
            }
        }

        public int Height
        {
            get
            {
                ThrowIfDisposed(); 
                return InvokeMemberValue(() => Item.Height);
            }
        }

        public WindowState WindowState
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => (WindowState)Item.WindowState);
            }
        }

        public WindowKind Type
        {
            get
            {
                ThrowIfDisposed(); 
                return (WindowKind)InvokeMemberValue(() => Item.Type);
            }
        }

        public LinkedWindows LinkedWindows
        {
            get
            {
                ThrowIfDisposed();
                return new LinkedWindows(InvokeMemberValue(() => Item.LinkedWindows));
            }
        }

        public Window LinkedWindowFrame
        {
            get
            {
                ThrowIfDisposed(); 
                return new Window(InvokeMemberValue(() => Item.LinkedWindowFrame));
            }
        }

        public int HWnd
        {
            get
            {
                ThrowIfDisposed(); 
                return InvokeMemberValue(() => Item.HWnd);
            }
        }
    }
}