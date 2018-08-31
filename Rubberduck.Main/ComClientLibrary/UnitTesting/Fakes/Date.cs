using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Date : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcGetDateVar");

        public Date()
        {
            InjectDelegate(new DateDelegate(DateCallback), ProcessAddress);
        }

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcGetDateVar(out object retVal);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void DateDelegate(IntPtr retVal);

        public void DateCallback(IntPtr retVal)
        {
            OnCallBack(true);
            if (!TrySetReturnValue())                          // specific invocation
            {
                TrySetReturnValue(true);                       // any invocation
            }
            if (PassThrough)
            {
                object result;
                rtcGetDateVar(out result);
                Marshal.GetNativeVariantForObject(result, retVal);
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
