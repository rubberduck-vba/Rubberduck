using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    public sealed class RawKeyboard : IRawDevice
    {
        public RawKeyboard(IntPtr hwnd, bool captureOnlyInForeground)
        {
            var rid = new RawInputDevice[1];
            rid[0].UsagePage = HidUsagePage.GENERIC;
            rid[0].Usage = HidUsage.Keyboard;
            rid[0].Flags = (captureOnlyInForeground ? RawInputDeviceFlags.NONE : RawInputDeviceFlags.INPUTSINK);
            rid[0].Target = hwnd;
            if (!User32.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(rid[0])))
            {
                throw new ApplicationException("Failed to register raw input device(s).");
            }
        }

        public event EventHandler<RawKeyEventArgs> RawKeyInputReceived;

        public void ProcessRawInput(InputData _rawBuffer)
        {
            if (_rawBuffer.header.dwType != (uint)DeviceType.RIM_TYPE_KEYBOARD)
            {
                return;
            }
            int virtualKey = _rawBuffer.data.keyboard.VKey;
            int makeCode = _rawBuffer.data.keyboard.Makecode;
            int flags = _rawBuffer.data.keyboard.Flags;
            if (virtualKey == Win32.KEYBOARD_OVERRUN_MAKE_CODE)
            {
                return;
            }
            var isE0BitSet = ((flags & Win32.RI_KEY_E0) != 0);
            var isBreakBitSet = ((flags & Win32.RI_KEY_BREAK) != 0);
            var args = new RawKeyEventArgs(
                            virtualKey,
                            KeyMap.GetKeyName(VirtualKeyCorrection(virtualKey, isE0BitSet, makeCode, _rawBuffer)).ToUpper(),
                            (WM)_rawBuffer.data.keyboard.Message,
                            isBreakBitSet ? "BREAK" : "MAKE");
            if (RawKeyInputReceived != null)
            {
                RawKeyInputReceived(this, args);
            }
        }

        private int VirtualKeyCorrection(int virtualKey, bool isE0BitSet, int makeCode, InputData _rawBuffer)
        {
            var correctedVKey = virtualKey;

            if (_rawBuffer.header.hDevice == IntPtr.Zero)
            {
                // When hDevice is 0 and the vkey is VK_CONTROL indicates the ZOOM key
                if (_rawBuffer.data.keyboard.VKey == Win32.VK_CONTROL)
                {
                    correctedVKey = Win32.VK_ZOOM;
                }
            }
            else
            {
                switch (virtualKey)
                {
                    // Right-hand CTRL and ALT have their e0 bit set 
                    case Win32.VK_CONTROL:
                        correctedVKey = isE0BitSet ? Win32.VK_RCONTROL : Win32.VK_LCONTROL;
                        break;
                    case Win32.VK_MENU:
                        correctedVKey = isE0BitSet ? Win32.VK_RMENU : Win32.VK_LMENU;
                        break;
                    case Win32.VK_SHIFT:
                        correctedVKey = makeCode == Win32.SC_SHIFT_R ? Win32.VK_RSHIFT : Win32.VK_LSHIFT;
                        break;
                    default:
                        correctedVKey = virtualKey;
                        break;
                }
            }

            return correctedVKey;
        }
    }
}
