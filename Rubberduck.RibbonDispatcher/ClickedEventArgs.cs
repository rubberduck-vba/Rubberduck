using System;

namespace Rubberduck.RibbonDispatcher
{
    public class ClickedEventArgs : EventArgs
    {
        public ClickedEventArgs(bool isPressed) { IsPressed = isPressed; }
        public bool IsPressed   { get; }
    }
}
