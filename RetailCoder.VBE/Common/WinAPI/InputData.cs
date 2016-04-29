using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    [StructLayout(LayoutKind.Sequential)]
    public struct InputData
    {
        public RawInputHeader header;           // 64 bit header size: 24  32 bit the header size: 16
        public RawData data;                    // Creating the rest in a struct allows the header size to align correctly for 32/64 bit
    }
}
