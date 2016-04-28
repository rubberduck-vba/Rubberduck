using System;

namespace RawInput_dll
{
    public interface IRawDevice
    {
        void ProcessRawInput(IntPtr hdevice);
        void EnumerateDevices();
    }
}
