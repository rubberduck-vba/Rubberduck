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
            if (!TrySetReturnValue())
            {
                TrySetReturnValue(true);
            }
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<TimeDelegate>(NativeFunctionAddress);
                nativeCall(retVal);
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
