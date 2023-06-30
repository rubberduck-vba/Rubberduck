using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class IMEStatus : FakeBase
    {
        public IMEStatus()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcIMEStatus");

            InjectDelegate(new IMEStatusDelegate(IMEStatusCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I2)]
        private delegate short IMEStatusDelegate();

        public short IMEStatusCallback()
        {
            OnCallBack(true);
            if (!TrySetReturnValue())
            {
                TrySetReturnValue(true);
            }

            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<IMEStatusDelegate>(NativeFunctionAddress);
                return nativeCall();
            }
            return Convert.ToInt16(ReturnValue ?? 0);
        } 
    }
}
