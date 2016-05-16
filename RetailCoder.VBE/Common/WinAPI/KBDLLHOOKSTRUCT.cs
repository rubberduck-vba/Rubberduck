using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    /// <remarks>
    /// This is named "struct" to be consistent with the Windows API name, but it is defined as a class since it is passed as a pointer (AKA reference) in SetWindowsHookEx and CallNextHookEx. 
    /// Or it can be defined as a struct and passed with ref.
    /// </remarks>>
    [StructLayout(LayoutKind.Sequential)]
    public class KBDLLHOOKSTRUCT
    {
        public uint vkCode;
        public uint scanCode;
        public KBDLLHOOKSTRUCTFlags flags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }
}