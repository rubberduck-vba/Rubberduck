using System;

namespace Rubberduck.Common.WinAPI
{
    public sealed class RawMouseEventArgs : EventArgs
    {
        private uint _message;
        private UsButtonFlags _ulButtons;

        public RawMouseEventArgs(
            uint message,
            UsButtonFlags ulButtons)
        {
            _message = message;
            _ulButtons = ulButtons;
        }

        public uint Message
        {
            get
            {
                return _message;
            }

            set
            {
                _message = value;
            }
        }

        public UsButtonFlags UlButtons
        {
            get
            {
                return _ulButtons;
            }

            set
            {
                _ulButtons = value;
            }
        }
    }
}
                                         

