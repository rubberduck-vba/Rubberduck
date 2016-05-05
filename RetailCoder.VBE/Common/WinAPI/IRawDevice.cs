using System;

namespace Rubberduck.Common.WinAPI
{
    public interface IRawDevice
    {
        void ProcessRawInput(IntPtr hdevice);
        void EnumerateDevices();
    }
}
