using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    public sealed class RawMouse : IRawDevice
    {
        public RawMouse(IntPtr hwnd, bool captureOnlyInForeground)
        {
            var rid = new RawInputDevice[1];
            rid[0].UsagePage = HidUsagePage.GENERIC;
            rid[0].Usage = HidUsage.Mouse;
            rid[0].Flags = (captureOnlyInForeground ? RawInputDeviceFlags.NONE : RawInputDeviceFlags.INPUTSINK);
            rid[0].Target = hwnd;
            if (!User32.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(rid[0])))
            {
                throw new ApplicationException("Failed to register raw input device(s).");
            }
        }

        public event EventHandler<RawMouseEventArgs> RawMouseInputReceived;

        public void ProcessRawInput(InputData _rawBuffer)
        {
            if (_rawBuffer.header.dwType != (uint)DeviceType.RIM_TYPE_MOUSE)
            {
                return;
            }
            var args = new RawMouseEventArgs(
                            _rawBuffer.data.keyboard.Message,
                            (UsButtonFlags)_rawBuffer.data.mouse.ulButtons);

            if (RawMouseInputReceived != null)
            {
                RawMouseInputReceived(this, args);
            }
        }
    }
}
