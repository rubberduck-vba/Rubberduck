using System;
using System.Diagnostics;
using System.Globalization;
using System.ComponentModel;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace RawInput_dll
{
    public sealed class RawMouse : IRawDevice
    {
        private readonly Dictionary<IntPtr, MouseClickEvent> _deviceList = new Dictionary<IntPtr, MouseClickEvent>();
        public delegate void DeviceEventHandler(object sender, MouseClickEventArgs e);
        public event DeviceEventHandler OnMouseClick;
        readonly object _padLock = new object();
        public int NumberOfKeyboards { get; private set; }
        static InputData _rawBuffer;

        public RawMouse(IntPtr hwnd, bool captureOnlyInForeground)
        {
            var rid = new RawInputDevice[1];
            rid[0].UsagePage = HidUsagePage.GENERIC;
            rid[0].Usage = HidUsage.Mouse;
            rid[0].Flags = (captureOnlyInForeground ? RawInputDeviceFlags.NONE : RawInputDeviceFlags.INPUTSINK) | RawInputDeviceFlags.DEVNOTIFY;
            rid[0].Target = hwnd;
            if (!Win32.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(rid[0])))
            {
                throw new ApplicationException("Failed to register raw input device(s).");
            }
        }

        public void EnumerateDevices()
        {
            lock (_padLock)
            {
                _deviceList.Clear();

                var keyboardNumber = 0;

                var globalDevice = new MouseClickEvent
                {
                    DeviceName = "Global Keyboard",
                    DeviceHandle = IntPtr.Zero,
                    DeviceType = Win32.GetDeviceType(DeviceType.RimTypeMouse),
                    Name = "Fake Keyboard. Some keys (ZOOM, MUTE, VOLUMEUP, VOLUMEDOWN) are sent to rawinput with a handle of zero.",
                    Source = keyboardNumber++.ToString(CultureInfo.InvariantCulture)
                };

                _deviceList.Add(globalDevice.DeviceHandle, globalDevice);

                var numberOfDevices = 0;
                uint deviceCount = 0;
                var dwSize = (Marshal.SizeOf(typeof(Rawinputdevicelist)));

                if (Win32.GetRawInputDeviceList(IntPtr.Zero, ref deviceCount, (uint)dwSize) == 0)
                {
                    var pRawInputDeviceList = Marshal.AllocHGlobal((int)(dwSize * deviceCount));
                    Win32.GetRawInputDeviceList(pRawInputDeviceList, ref deviceCount, (uint)dwSize);

                    for (var i = 0; i < deviceCount; i++)
                    {
                        uint pcbSize = 0;

                        // On Window 8 64bit when compiling against .Net > 3.5 using .ToInt32 you will generate an arithmetic overflow. Leave as it is for 32bit/64bit applications
                        var rid = (Rawinputdevicelist)Marshal.PtrToStructure(new IntPtr((pRawInputDeviceList.ToInt64() + (dwSize * i))), typeof(Rawinputdevicelist));

                        Win32.GetRawInputDeviceInfo(rid.hDevice, RawInputDeviceInfo.RIDI_DEVICENAME, IntPtr.Zero, ref pcbSize);

                        if (pcbSize <= 0) continue;

                        var pData = Marshal.AllocHGlobal((int)pcbSize);
                        Win32.GetRawInputDeviceInfo(rid.hDevice, RawInputDeviceInfo.RIDI_DEVICENAME, pData, ref pcbSize);
                        var deviceName = Marshal.PtrToStringAnsi(pData);

                        if (rid.dwType == DeviceType.RimTypeMouse || rid.dwType == DeviceType.RimTypeHid)
                        {
                            var deviceDesc = Win32.GetDeviceDescription(deviceName);

                            var dInfo = new MouseClickEvent
                            {
                                DeviceName = Marshal.PtrToStringAnsi(pData),
                                DeviceHandle = rid.hDevice,
                                DeviceType = Win32.GetDeviceType(rid.dwType),
                                Name = deviceDesc,
                                Source = keyboardNumber++.ToString(CultureInfo.InvariantCulture)
                            };

                            if (!_deviceList.ContainsKey(rid.hDevice))
                            {
                                numberOfDevices++;
                                _deviceList.Add(rid.hDevice, dInfo);
                            }
                        }

                        Marshal.FreeHGlobal(pData);
                    }

                    Marshal.FreeHGlobal(pRawInputDeviceList);

                    NumberOfKeyboards = numberOfDevices;
                    Debug.WriteLine("EnumerateDevices() found {0} Keyboard(s)", NumberOfKeyboards);
                    return;
                }
            }

            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        public void ProcessRawInput(IntPtr hdevice)
        {
            if (_deviceList.Count == 0) return;

            var dwSize = 0;
            Win32.GetRawInputData(hdevice, DataCommand.RID_INPUT, IntPtr.Zero, ref dwSize, Marshal.SizeOf(typeof(Rawinputheader)));

            if (dwSize != Win32.GetRawInputData(hdevice, DataCommand.RID_INPUT, out _rawBuffer, ref dwSize, Marshal.SizeOf(typeof(Rawinputheader))))
            {
                Debug.WriteLine("Error getting the rawinput buffer");
                return;
            }

            MouseClickEvent mouseClickEvent;

            if (_deviceList.ContainsKey(_rawBuffer.header.hDevice))
            {
                lock (_padLock)
                {
                    mouseClickEvent = _deviceList[_rawBuffer.header.hDevice];
                }
            }
            else
            {
                return;
            }

            mouseClickEvent.Message = _rawBuffer.data.keyboard.Message;
            mouseClickEvent.ulButtons =   (UsButtonFlags)_rawBuffer.data.mouse.ulButtons;

            if (_rawBuffer.data.mouse.ulButtons != 0)
            {
                Debug.WriteLine(mouseClickEvent.ulButtons);
            }

            if (OnMouseClick != null)
            {
                OnMouseClick(this, new MouseClickEventArgs(mouseClickEvent));
            }
        }
    }
}
