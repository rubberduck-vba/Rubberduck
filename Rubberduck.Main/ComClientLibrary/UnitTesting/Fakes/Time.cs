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
                FakesProvider.SuspendFake(typeof(Now));
                var nativeCall = Marshal.GetDelegateForFunctionPointer<TimeDelegate>(NativeFunctionAddress);
                nativeCall(retVal);
                FakesProvider.ResumeFake(typeof(Now));
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
