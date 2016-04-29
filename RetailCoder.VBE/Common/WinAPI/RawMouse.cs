using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    public sealed class RawMouse : RawDevice
    {
        private InputData _rawBuffer;

        public RawMouse(IntPtr hwnd, bool captureOnlyInForeground)
        {
            var rid = new RawInputDevice[1];
            rid[0].UsagePage = HidUsagePage.GENERIC;
            rid[0].Usage = HidUsage.Mouse;
            rid[0].Flags = (captureOnlyInForeground ? RawInputDeviceFlags.NONE : RawInputDeviceFlags.INPUTSINK) | RawInputDeviceFlags.DEVNOTIFY;
            rid[0].Target = hwnd;
            if (!User32.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(rid[0])))
            {
                throw new ApplicationException("Failed to register raw input device(s).");
            }
        }

        public event EventHandler<RawMouseEventArgs> RawMouseInputReceived;

        public override void EnumerateDevices()
        {
            EnumerateDevices(DeviceType.RIM_TYPE_MOUSE);
        }

        public override void ProcessRawInput(IntPtr hdevice)
        {
            if (DeviceList.Count == 0) return;
            var dwSize = 0;
            User32.GetRawInputData(hdevice, DataCommand.RID_INPUT, IntPtr.Zero, ref dwSize, Marshal.SizeOf(typeof(RawInputHeader)));
            if (dwSize != User32.GetRawInputData(hdevice, DataCommand.RID_INPUT, out _rawBuffer, ref dwSize, Marshal.SizeOf(typeof(RawInputHeader))))
            {
                Debug.WriteLine("Error getting the rawinput buffer");
                return;
            }
            EnumeratedDevice enumeratedDevice;
            if (DeviceList.ContainsKey(_rawBuffer.header.hDevice))
            {
                lock (PadLock)
                {
                    enumeratedDevice = DeviceList[_rawBuffer.header.hDevice];
                }
            }
            else
            {
                return;
            }
            
            var args = new RawMouseEventArgs(
                            enumeratedDevice.DeviceName,
                            enumeratedDevice.DeviceType,
                            enumeratedDevice.DeviceHandle,
                            enumeratedDevice.Name,
                            enumeratedDevice.Source,
                            _rawBuffer.data.keyboard.Message,
                            (UsButtonFlags)_rawBuffer.data.mouse.ulButtons);

            if (RawMouseInputReceived != null)
            {
                RawMouseInputReceived(this, args);
            }
        }
    }
}
