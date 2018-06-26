using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Now : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcGetPresentDate");

        public Now()
        {
            InjectDelegate(new NowDelegate(NowCallback), ProcessAddress);
        }

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcGetPresentDate(out object retVal);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void NowDelegate(IntPtr retVal);

        public void NowCallback(IntPtr retVal)
        {
            OnCallBack(true);
            if (!TrySetReturnValue())                          // specific invocation
            {
                TrySetReturnValue(true);                       // any invocation
            }
            if (PassThrough)
            {
                object result;
                rtcGetPresentDate(out result);
                Marshal.GetNativeVariantForObject(result, retVal);
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
