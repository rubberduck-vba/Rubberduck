using System;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IWindow : ISafeComWrapper, IEquatable<IWindow>
    {
        int HWnd { get; }
        string Caption { get; }
        bool IsVisible { get; set; }
        int Left { get; set; }
        int Top { get; set; }
        int Width { get; set; }
        int Height { get; set; }
        WindowState WindowState { get; }
        WindowKind Type { get; }
        IVBE VBE { get; }
        IWindow LinkedWindowFrame { get; }
        IWindows Collection { get; }
        ILinkedWindows LinkedWindows { get; }
        IntPtr Handle();
        void Close();
        void SetFocus();
        void SetKind(WindowKind eKind);
        void Detach();
        void Attach(int lWindowHandle);
    }
}