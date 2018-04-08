using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
        public IWindow Window { get; }
        public ICodePane CodePane { get; }
        public FocusType EventType { get; }

        public WindowChangedEventArgs(IntPtr hwnd, IWindow window, ICodePane pane, FocusType type)
        {
            Hwnd = hwnd;
            Window = window;
            CodePane = pane;
            EventType = type;
        }
    }
}
