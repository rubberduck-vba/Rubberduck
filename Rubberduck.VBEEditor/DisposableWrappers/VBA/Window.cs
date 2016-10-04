using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers.VBA
{
    public class Window : SafeComWrapper<Microsoft.Vbe.Interop.Window>
    {
        public Window(Microsoft.Vbe.Interop.Window window)
            : base(window)
        {
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

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public Windows Collection { get { return new Windows(InvokeResult(() => ComObject.Collection)); } }

        public string Caption { get { return InvokeResult(() => ComObject.Caption); } }

        public bool Visible
        {
            get { return InvokeResult(() => ComObject.Visible); }
            set { Invoke(() => ComObject.Visible = value); }
        }

        public int Left
        {
            get { ThrowIfDisposed(); return InvokeResult(() => ComObject.Left); }
            set { ThrowIfDisposed(); Invoke(() => ComObject.Left = value); }
        }

        public int Top
        {
            get { ThrowIfDisposed(); return InvokeResult(() => ComObject.Top); }
            set { ThrowIfDisposed(); Invoke(() => ComObject.Top = value); }
        }

        public int Width
        {
            get { ThrowIfDisposed(); return InvokeResult(() => ComObject.Width); }
            set { ThrowIfDisposed(); Invoke(() => ComObject.Width = value); }
        }

        public int Height
        {
            get { ThrowIfDisposed(); return InvokeResult(() => ComObject.Height); }
            set { ThrowIfDisposed(); Invoke(() => ComObject.Height = value); }
        }

        public WindowState WindowState { get { return InvokeResult(() => (WindowState)ComObject.WindowState); } }

        public WindowKind Type { get { return (WindowKind)InvokeResult(() => ComObject.Type); } }

        public LinkedWindows LinkedWindows { get { return new LinkedWindows(InvokeResult(() => ComObject.LinkedWindows)); } }

        public Window LinkedWindowFrame { get { return new Window(InvokeResult(() => ComObject.LinkedWindowFrame)); } }

        public int HWnd { get { return InvokeResult(() => ComObject.HWnd); } }
    }
}