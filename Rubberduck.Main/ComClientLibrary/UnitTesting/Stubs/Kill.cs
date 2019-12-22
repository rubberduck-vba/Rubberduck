using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    internal class Kill : StubBase
    {
        public Kill()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcKillFiles");

            InjectDelegate(new KillDelegate(KillCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void KillDelegate(IntPtr pathname);

        public void KillCallback(IntPtr pathname)
        {
            OnCallBack(true);

            TrackUsage("pathname", pathname);
            if (PassThrough)
            {
                VbeProvider.VbeNativeApi.KillFiles(pathname);
            }
        }
    }
}
