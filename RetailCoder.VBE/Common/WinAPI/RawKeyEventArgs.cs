using System;

namespace Rubberduck.Common.WinAPI
{
    public sealed class RawKeyEventArgs : EventArgs
    {
        private int _vKey;                  // Virtual Key. Corrected for L/R keys(i.e. LSHIFT/RSHIFT) and Zoom
        private string _vKeyName;           // Virtual Key Name. Corrected for L/R keys(i.e. LSHIFT/RSHIFT) and Zoom
        private WM _message;                // WM_KEYDOWN or WM_KEYUP        
        private string _keyPressState;      // MAKE or BREAK

        public RawKeyEventArgs(
            int vKey,
            string vKeyName,
            WM message,
            string keyPressState)
        {
            _vKey = vKey;
            _vKeyName = vKeyName;
            _message = message;
            _keyPressState = keyPressState;
        }

        public string KeyPressState
        {
            get
            {
                return _keyPressState;
            }

            set
            {
                _keyPressState = value;
            }
        }

        public WM Message
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

        public string VKeyName
        {
            get
            {
                return _vKeyName;
            }

            set
            {
                _vKeyName = value;
            }
        }

        public int VKey
        {
            get
            {
                return _vKey;
            }

            set
            {
                _vKey = value;
            }
        }
    }
}
                                         

