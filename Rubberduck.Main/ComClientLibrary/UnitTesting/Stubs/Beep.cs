using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    internal class Beep : StubBase
    {
        public Beep()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeRuntime.DllName, "rtcBeep");

            InjectDelegate(new BeepDelegate(BeepCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void BeepDelegate();

        public void BeepCallback()
        {
            OnCallBack(true);

            if (PassThrough)
            {
                VbeProvider.VbeRuntime.Beep();
            }
        }
    }
}
