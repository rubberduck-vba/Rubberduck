using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    public abstract class RawDevice : IRawDevice
    {
        private readonly Dictionary<IntPtr, EnumeratedDevice> _deviceList = new Dictionary<IntPtr, EnumeratedDevice>();
        private readonly object _padLock = new object();

        protected Dictionary<IntPtr, EnumeratedDevice> DeviceList
        {
            get
            {
                return _deviceList;
            }
        }

        protected object PadLock
        {
            get
            {
                return _padLock;
            }
        }

        protected void EnumerateDevices(DeviceType deviceType)
        {
            lock (_padLock)
            {
                _deviceList.Clear();
                var deviceNumber = 0;
                var globalDevice = new EnumeratedDevice
                {
                    DeviceName = "Global Device",
                    DeviceHandle = IntPtr.Zero,
                    DeviceType = Win32.GetDeviceType(deviceType),
                    Name = "Fake Device",
                    Source = deviceNumber++.ToString(CultureInfo.InvariantCulture)
                };

                _deviceList.Add(globalDevice.DeviceHandle, globalDevice);

                uint deviceCount = 0;
                var dwSize = (Marshal.SizeOf(typeof(RawInputDeviceList)));
                if (User32.GetRawInputDeviceList(IntPtr.Zero, ref deviceCount, (uint)dwSize) == 0)
                {
                    var pRawInputDeviceList = Marshal.AllocHGlobal((int)(dwSize * deviceCount));
                    User32.GetRawInputDeviceList(pRawInputDeviceList, ref deviceCount, (uint)dwSize);
                    for (var i = 0; i < deviceCount; i++)
                    {
                        uint pcbSize = 0;

                        // On Window 8 64bit when compiling against .Net > 3.5 using .ToInt32 you will generate an arithmetic overflow. Leave as it is for 32bit/64bit applications
                        var rid = (RawInputDeviceList)Marshal.PtrToStructure(new IntPtr((pRawInputDeviceList.ToInt64() + (dwSize * i))), typeof(RawInputDeviceList));

                        User32.GetRawInputDeviceInfo(rid.hDevice, RawInputDeviceInfo.RIDI_DEVICENAME, IntPtr.Zero, ref pcbSize);

                        if (pcbSize <= 0) continue;

                        var pData = Marshal.AllocHGlobal((int)pcbSize);
                        User32.GetRawInputDeviceInfo(rid.hDevice, RawInputDeviceInfo.RIDI_DEVICENAME, pData, ref pcbSize);
                        var deviceName = Marshal.PtrToStringAnsi(pData);

                        if (rid.dwType == (uint)deviceType || rid.dwType == (uint)DeviceType.RIM_TYPE_HID)
                        {
                            var deviceDesc = Win32.GetDeviceDescription(deviceName);

                            var dInfo = new EnumeratedDevice
                            {
                                DeviceName = Marshal.PtrToStringAnsi(pData),
                                DeviceHandle = rid.hDevice,
                                DeviceType = Win32.GetDeviceType((DeviceType)rid.dwType),
                                Name = deviceDesc,
                                Source = deviceNumber++.ToString(CultureInfo.InvariantCulture)
                            };

                            if (!_deviceList.ContainsKey(rid.hDevice))
                            {
                                _deviceList.Add(rid.hDevice, dInfo);
                            }
                        }

                        Marshal.FreeHGlobal(pData);
                    }
                    Marshal.FreeHGlobal(pRawInputDeviceList);
                    return;
                }
            }
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        public abstract void ProcessRawInput(IntPtr hdevice);
        public abstract void EnumerateDevices();
    }
}
