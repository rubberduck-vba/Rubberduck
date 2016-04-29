using System;

namespace Rubberduck.Common.WinAPI
{
    public sealed class RawMouseEventArgs : EventArgs
    {
        private string _deviceName;       // i.e. \\?\HID#VID_045E&PID_00DD&MI_00#8&1eb402&0&0000#{884b96c3-56ef-11d1-bc8c-00a0c91405dd}
        private string _deviceType;       // MOUSE or HID
        private IntPtr _deviceHandle;     // Handle to the device that send the input
        private string _name;
        private string _source;
        private uint _message;
        private UsButtonFlags _ulButtons;

        public RawMouseEventArgs(
            string deviceName,
            string deviceType,
            IntPtr deviceHandle,
            string name,
            string source,
            uint message,
            UsButtonFlags ulButtons)
        {
            _deviceName = deviceName;
            _deviceType = deviceType;
            _deviceHandle = deviceHandle;
            _name = name;
            _source = source;
            _message = message;
            _ulButtons = ulButtons;
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

        public override string ToString()
        {
            return string.Format("Device\n DeviceName: {0}\n DeviceType: {1}\n DeviceHandle: {2}\n Name: {3}\n", _deviceName, _deviceType, _deviceHandle.ToInt64().ToString("X"), _name);
        }
    }
}
                                         

