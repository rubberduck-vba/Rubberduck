using System;

namespace Rubberduck.VBEditor.Events
{
    public enum FocusType
    {
        GotFocus,
        LostFocus,
        ChildFocus
    }

    public class WindowChangedEventArgs : EventArgs
    {
        public IntPtr Hwnd { get; }
        public FocusType EventType { get; }

        public WindowChangedEventArgs(IntPtr hwnd, FocusType type)
        {
            Hwnd = hwnd;
            EventType = type;
        }
    }
}
