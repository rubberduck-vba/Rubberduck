using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Time : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcGetTimeVar");

        public Time()
        {
            InjectDelegate(new TimeDelegate(TimeCallback), ProcessAddress);
        }

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcGetTimeVar(out object retVal);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void TimeDelegate(IntPtr retVal);

        public void TimeCallback(IntPtr retVal)
        {
            OnCallBack(true);
            if (!TrySetReturnValue())                          // specific invocation
            {
                TrySetReturnValue(true);                       // any invocation
            }
            if (PassThrough)
            {
                object result;
                rtcGetTimeVar(out result);
                Marshal.GetNativeVariantForObject(result, retVal);
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
