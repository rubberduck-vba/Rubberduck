using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    internal class Kill : StubBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcKillFiles");

        public Kill()
        {
            InjectDelegate(new KillDelegate(KillCallback), ProcessAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void KillDelegate(IntPtr pathname);

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcKillFiles(IntPtr pathname);

        public void KillCallback(IntPtr pathname)
        {
            OnCallBack(true);

            TrackUsage("pathname", pathname);
            if (PassThrough)
            {
                rtcKillFiles(pathname);
            }
        }
    }
}
