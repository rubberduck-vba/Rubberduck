using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Rubberduck.Common.WinAPI
{
    public class RawInput : NativeWindow
    {
        static readonly Guid DeviceInterfaceHid = new Guid("4D1E55B2-F16F-11CF-88CB-001111000030");
        private readonly List<IRawDevice> _devices;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RawInput(IntPtr parentHandle)
        {
            AssignHandle(parentHandle);
            _devices = new List<IRawDevice>();
        }

        public void AddDevice(IRawDevice device)
        {
            _devices.Add(device);
        }

        public IRawDevice CreateKeyboard()
        {
            return new RawKeyboard(Handle, true);
        }

        public IRawDevice CreateMouse()
        {
            return new RawMouse(Handle, true);
        }        

        protected override void WndProc(ref Message message)
        {
            switch ((WM)message.Msg)
            {
                case WM.INPUT:
                    {
                        InputData _rawBuffer;
                        var dwSize = 0;
                        User32.GetRawInputData(message.LParam, DataCommand.RID_INPUT, IntPtr.Zero, ref dwSize, Marshal.SizeOf(typeof(RawInputHeader)));
                        int res = User32.GetRawInputData(message.LParam, DataCommand.RID_INPUT, out _rawBuffer, ref dwSize, Marshal.SizeOf(typeof(RawInputHeader)));
                        if (dwSize != res)
                        {
                            var ex = new Win32Exception(Marshal.GetLastWin32Error());
                            _logger.Error(ex, "Error getting the rawinput buffer: {0}", ex.Message);
                            return;
                        }
                        foreach (var device in _devices)
                        {
                            device.ProcessRawInput(_rawBuffer);
                        }
                    }
                    break;
            }
            base.WndProc(ref message);
        }
    }
}
