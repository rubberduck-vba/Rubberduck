using System;

namespace Rubberduck.Common.WinAPI
{
    public sealed class EnumeratedDevice
    {
        public string DeviceName { get; set; }
        public IntPtr DeviceHandle { get; set; }
        public string DeviceType { get; set; }
        public string Name { get; set; }
        public string Source { get; set; }
    }
}
