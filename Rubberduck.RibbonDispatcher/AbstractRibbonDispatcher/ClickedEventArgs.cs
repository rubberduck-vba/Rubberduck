using System;

namespace AbstractRibbonDispatcher
{
    public class ClickedEventArgs : EventArgs
    {
        public ClickedEventArgs(bool isPressed) { IsPressed = isPressed; }
        public bool IsPressed   { get; }
    }
}
