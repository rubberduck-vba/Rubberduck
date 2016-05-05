using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct RawInputDeviceList
    {
        public IntPtr hDevice;
        public uint dwType;
    }
}
