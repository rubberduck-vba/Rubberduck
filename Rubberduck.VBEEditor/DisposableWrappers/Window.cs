using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.DisposableWrappers
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
            Invoke(kind => ComObject.SetKind((vbext_WindowType)kind), eKind);
        }

        public void Detach()
        {
            Invoke(() => ComObject.Detach());
        }

        public void Attach(int lWindowHandle)
        {
            Invoke(handle => ComObject.Attach(handle), lWindowHandle);
        }

        public VBE VBE { get { return new VBE(InvokeResult(() => ComObject.VBE)); } }

        public Windows Collection { get { return new Windows(InvokeResult(() => ComObject.Collection)); } }

        public string Caption { get { return InvokeResult(() => ComObject.Caption); } }

        public bool Visible { get { return InvokeResult(() => ComObject.Visible); } }

        public int Left { get { return InvokeResult(() => ComObject.Left); } }

        public int Top { get { return InvokeResult(() => ComObject.Top); } }

        public int Width { get { ThrowIfDisposed();  return InvokeResult(() => ComObject.Width); } }

        public int Height { get { return InvokeResult(() => ComObject.Height); } }

        public WindowState WindowState { get { return InvokeResult(() => (WindowState)ComObject.WindowState); } }

        public WindowKind Type { get { return (WindowKind)InvokeResult(() => ComObject.Type); } }

        public LinkedWindows LinkedWindows { get { return new LinkedWindows(InvokeResult(() => ComObject.LinkedWindows)); } }

        public Window LinkedWindowFrame { get { return new Window(InvokeResult(() => ComObject.LinkedWindowFrame)); } }

        public int HWnd { get { return InvokeResult(() => ComObject.HWnd); } }
    }
}