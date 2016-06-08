using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct RawKeyboardSTRUCT
    {
        public ushort Makecode;                 // Scan code from the key depression
        public ushort Flags;                    // One or more of RI_KEY_MAKE, RI_KEY_BREAK, RI_KEY_E0, RI_KEY_E1
        private readonly ushort Reserved;       // Always 0    
        public ushort VKey;                     // Virtual Key Code
        public uint Message;                    // Corresponding Windows message for exmaple (WM_KEYDOWN, WM_SYASKEYDOWN etc)
        public uint ExtraInformation;           // The device-specific addition information for the event (seems to always be zero for keyboards)

        public override string ToString()
        {
            return string.Format("Rawkeyboard\n Makecode: {0}\n Makecode(hex) : {0:X}\n Flags: {1}\n Reserved: {2}\n VKeyName: {3}\n Message: {4}\n ExtraInformation {5}\n",
                                                Makecode, Flags, Reserved, VKey, Message, ExtraInformation);
        }
    }
}
