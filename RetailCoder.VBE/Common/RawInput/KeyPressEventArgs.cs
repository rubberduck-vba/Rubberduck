using System;

namespace RawInput_dll
{
    public class KeyPressEventArgs : EventArgs
    {
        public KeyPressEventArgs(KeyPressEvent arg)
        {
            KeyPressEvent = arg;
        }
        
        public KeyPressEvent KeyPressEvent { get; private set; }
    }
}
