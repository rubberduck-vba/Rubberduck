using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class GetAttr : FakeBase
    {
        public GetAttr()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcGetFileAttr");

            InjectDelegate(new GetAttrDelegate(GetAttrCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I2)]
        private delegate short GetAttrDelegate(IntPtr pathname);

        public short GetAttrCallback(IntPtr pathname)
        {
            OnCallBack();

            var pathNameArg = Marshal.PtrToStringBSTR(pathname);
            TrackUsage("pathname", pathNameArg, Tokens.String);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<GetAttrDelegate>(NativeFunctionAddress);
                return nativeCall(pathname);
            }

            return Convert.ToInt16(ReturnValue ?? 0);
        }    
    }
}
