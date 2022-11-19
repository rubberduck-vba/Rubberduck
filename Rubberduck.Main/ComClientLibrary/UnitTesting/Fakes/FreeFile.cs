using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class FreeFile : FakeBase
    {
        public FreeFile()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcFreeFile");

            InjectDelegate(new FreeFileDelegate(FreeFileCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I2)]
        private delegate short FreeFileDelegate(IntPtr rangenumber);

        public short FreeFileCallback(IntPtr rangenumber)
        {
            OnCallBack();

            TrackUsage("rangenumber", rangenumber);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<FreeFileDelegate>(NativeFunctionAddress);
                return nativeCall(rangenumber);
            }

            return Convert.ToInt16(ReturnValue ?? 0);
        }    
    }
}
