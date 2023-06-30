using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Date : FakeBase
    {
        public Date()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcGetDateVar");

            InjectDelegate(new DateDelegate(DateCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void DateDelegate(IntPtr retVal);

        public void DateCallback(IntPtr retVal)
        {
            OnCallBack(true);
            if (!TrySetReturnValue())
            {
                TrySetReturnValue(true);
            }
            if (PassThrough)
            {
                FakesProvider.SuspendFake(typeof(Now));
                var nativeCall = Marshal.GetDelegateForFunctionPointer<DateDelegate>(NativeFunctionAddress);
                nativeCall(retVal);
                FakesProvider.ResumeFake(typeof(Now));
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
