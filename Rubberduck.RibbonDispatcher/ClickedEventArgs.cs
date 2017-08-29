using System;

namespace Rubberduck.RibbonDispatcher
{
    [CLSCompliant(true)]
    public class ClickedEventArgs : EventArgs
    {
        public ClickedEventArgs(bool isPressed) { IsPressed = isPressed; }
        public bool IsPressed   { get; }
    }
}
