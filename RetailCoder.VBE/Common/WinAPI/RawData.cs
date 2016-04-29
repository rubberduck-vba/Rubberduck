using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    [StructLayout(LayoutKind.Explicit)]
    public struct RawData
    {
        [FieldOffset(0)]
        internal RawMouseSTRUCT mouse;
        [FieldOffset(0)]
        internal RawKeyboardSTRUCT keyboard;
        [FieldOffset(0)]
        internal RawHID hid;
    }
}
