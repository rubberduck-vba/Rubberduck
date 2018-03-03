using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class ChDrive : StubBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcChangeDrive");

        public ChDrive()
        {
            InjectDelegate(new ChDriveDelegate(ChDriveCallback), ProcessAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void ChDriveDelegate(IntPtr driveletter);

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcChangeDrive(IntPtr driveletter);

        public void ChDriveCallback(IntPtr driveletter)
        {
            OnCallBack(true);

            var driveletterArg = Marshal.PtrToStringBSTR(driveletter);

            TrackUsage("driveletter", driveletterArg, Tokens.String);
            if (PassThrough)
            {
                rtcChangeDrive(driveletter);
            }
        }
    }
}
