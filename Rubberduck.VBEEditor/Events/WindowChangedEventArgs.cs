using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class WindowChangedEventArgs : EventArgs
    {
        public enum FocusType
        {
            GotFocus,
            LostFocus
        }

        public IntPtr Hwnd { get; private set; }
        public IWindow Window { get; private set; }
        public FocusType EventType { get; private set; }

        public WindowChangedEventArgs(IntPtr hwnd, IWindow window, FocusType type)
        {
            Hwnd = hwnd;
            Window = window;
            EventType = type;
        }
    }
}
