﻿using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class ChDrive : StubBase
    {
        public ChDrive()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeRuntime.DllName, "rtcChangeDrive");

            InjectDelegate(new ChDriveDelegate(ChDriveCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void ChDriveDelegate(IntPtr driveletter);

        public void ChDriveCallback(IntPtr driveletter)
        {
            OnCallBack(true);

            var driveletterArg = Marshal.PtrToStringBSTR(driveletter);

            TrackUsage("driveletter", driveletterArg, Tokens.String);
            if (PassThrough)
            {
                VbeProvider.VbeRuntime.ChangeDrive(driveletter);
            }
        }
    }
}
