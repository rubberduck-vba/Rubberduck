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
        public IntPtr Hwnd { get; private set; }
        public IWindow Window { get; private set; }
        public ICodePane CodePane { get; private set; }
        public FocusType EventType { get; private set; }

        public WindowChangedEventArgs(IntPtr hwnd, IWindow window, ICodePane pane, FocusType type)
        {
            Hwnd = hwnd;
            Window = window;
            CodePane = pane;
            EventType = type;
        }
    }
}
