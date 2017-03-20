using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    internal class Beep : StubBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcBeep");

        public Beep() : base (ProcessAddress)
        {
            InjectDelegate(new BeepDelegate(BeepCallback));
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void BeepDelegate();

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcBeep();

        public void BeepCallback()
        {
            OnCallBack(true);

            if (PassThrough)
            {
                rtcBeep();
            }
        }
    }
}
