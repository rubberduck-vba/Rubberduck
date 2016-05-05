using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct RawInputDevice
    {
        internal HidUsagePage UsagePage;
        internal HidUsage Usage;
        internal RawInputDeviceFlags Flags;
        internal IntPtr Target;

        public override string ToString()
        {
            return string.Format("{0}/{1}, flags: {2}, target: {3}", UsagePage, Usage, Flags, Target);
        }
    }
}
