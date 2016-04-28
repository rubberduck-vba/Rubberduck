using System;

namespace RawInput_dll
{
    public class MouseClickEventArgs : EventArgs
    {
        public MouseClickEventArgs(MouseClickEvent arg)
        {
            MouseClickEvent = arg;
        }
        
        public MouseClickEvent MouseClickEvent { get; private set; }
    }
}
