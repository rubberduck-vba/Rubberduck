using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Time : FakeBase
    {
        public Time()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcGetTimeVar");

            InjectDelegate(new TimeDelegate(TimeCallback), processAddress);
        }

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
                VbeProvider.VbeNativeApi.GetTimeVar(out result);
                Marshal.GetNativeVariantForObject(result, retVal);
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
