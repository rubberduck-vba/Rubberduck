using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace Rubberduck.VBEditor.Events
{
    public class WindowChangedEventArgs : EventArgs
    {
        public IntPtr Hwnd { get; private set; }
        public IWindow Window { get; private set; }

        public WindowChangedEventArgs(IntPtr hwnd, IWindow window)
        {
            Hwnd = hwnd;
            Window = window;
        }
    }
}
