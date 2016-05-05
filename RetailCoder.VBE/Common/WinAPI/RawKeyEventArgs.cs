using System;

namespace Rubberduck.Common.WinAPI
{
    public sealed class RawKeyEventArgs : EventArgs
    {
        private string _deviceName;         // i.e. \\?\HID#VID_045E&PID_00DD&MI_00#8&1eb402&0&0000#{884b96c3-56ef-11d1-bc8c-00a0c91405dd}
        private string _deviceType;         // KEYBOARD or HID
        private IntPtr _deviceHandle;       // Handle to the device that send the input
        private string _name;               // i.e. Microsoft USB Comfort Curve Keyboard 2000 (Mouse and Keyboard Center)
        private string _source;             // Keyboard_XX
        private int _vKey;                  // Virtual Key. Corrected for L/R keys(i.e. LSHIFT/RSHIFT) and Zoom
        private string _vKeyName;           // Virtual Key Name. Corrected for L/R keys(i.e. LSHIFT/RSHIFT) and Zoom
        private WM _message;              // WM_KEYDOWN or WM_KEYUP        
        private string _keyPressState;      // MAKE or BREAK

        public RawKeyEventArgs(
            string deviceName,
            string deviceType,
            IntPtr deviceHandle,
            string name,
            string source,
            int vKey,
            string vKeyName,
            WM message,
            string keyPressState)
        {
            _deviceName = deviceName;
            _deviceType = deviceType;
            _deviceHandle = deviceHandle;
            _name = name;
            _source = source;
            _vKey = vKey;
            _vKeyName = vKeyName;
            _message = message;
            _keyPressState = keyPressState;
        }

        public string Source
        {
            get
            {
                return _source;
            }
            set
            {
                _source = string.Format("Keyboard_{0}", value.PadLeft(2, '0'));
            }
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

        public string Name
        {
            get
            {
                return _name;
            }

            set
            {
                _name = value;
            }
        }

        public IntPtr DeviceHandle
        {
            get
            {
                return _deviceHandle;
            }

            set
            {
                _deviceHandle = value;
            }
        }

        public string DeviceType
        {
            get
            {
                return _deviceType;
            }

            set
            {
                _deviceType = value;
            }
        }

        public string DeviceName
        {
            get
            {
                return _deviceName;
            }

            set
            {
                _deviceName = value;
            }
        }

        public override string ToString()
        {
            return string.Format("Device\n DeviceName: {0}\n DeviceType: {1}\n DeviceHandle: {2}\n Name: {3}\n", _deviceName, _deviceType, _deviceHandle.ToInt64().ToString("X"), _name);
        }
    }
}
                                         

