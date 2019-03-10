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
            if (!TrySetReturnValue())                          // specific invocation
            {
                TrySetReturnValue(true);                       // any invocation
            }
            if (PassThrough)
            {
                object result;
                VbeProvider.VbeNativeApi.GetDateVar(out result);
                Marshal.GetNativeVariantForObject(result, retVal);
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
