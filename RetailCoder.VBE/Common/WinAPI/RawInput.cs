using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.UI;

namespace Rubberduck.Common.WinAPI
{
    public class RawInput : SubclassingWindow
    {
        static readonly Guid DeviceInterfaceHid = new Guid("4D1E55B2-F16F-11CF-88CB-001111000030");
        private readonly List<IRawDevice> _devices;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static int _currentId;

        public RawInput(IntPtr parentHandle) : base (new IntPtr((int)parentHandle + (_currentId++)), parentHandle)
        {
            _devices = new List<IRawDevice>();
        }

        public void AddDevice(IRawDevice device)
        {
            _devices.Add(device);
        }

        public IRawDevice CreateKeyboard()
        {
            return new RawKeyboard(Hwnd, true);
        }

        public IRawDevice CreateMouse()
        {
            return new RawMouse(Hwnd, true);
        }

        public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            switch ((WM)msg)
            {
                case WM.INPUT:
                    {
                        if (lParam == IntPtr.Zero)
                        {
                            break;
                        }
                        InputData rawBuffer;
                        var dwSize = 0;
                        var res = User32.GetRawInputData(lParam, DataCommand.RID_INPUT, IntPtr.Zero, ref dwSize, Marshal.SizeOf(typeof(RawInputHeader)));
                        if (res != 0)
                        {
                            var ex = new Win32Exception(Marshal.GetLastWin32Error());
                            Logger.Error(ex, "Error sizing the rawinput buffer: {0}", ex.Message);
                            break;
                        }

                        res = User32.GetRawInputData(lParam, DataCommand.RID_INPUT, out rawBuffer, ref dwSize, Marshal.SizeOf(typeof(RawInputHeader)));
                        if (res == -1)
                        {
                            var ex = new Win32Exception(Marshal.GetLastWin32Error());
                            Logger.Error(ex, "Error getting the rawinput buffer: {0}", ex.Message);
                            break;
                        }
                        if (res == dwSize)
                        {
                            foreach (var device in _devices)
                            {
                                device.ProcessRawInput(rawBuffer);
                            }
                        }
                        else
                        {
                            //Something is seriously f'd up with Windows - the number of bytes copied does not match the reported buffer size.
                        }
                    }
                    break;
            }
            return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
        }
    }
}
