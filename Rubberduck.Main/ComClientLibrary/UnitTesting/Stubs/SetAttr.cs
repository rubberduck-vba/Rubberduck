using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class SetAttr : StubBase
    {
        public SetAttr()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcSetFileAttr");

            InjectDelegate(new SetAttrDelegate(SetAttrCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void SetAttrDelegate(IntPtr pathname, short attributes);

        public void SetAttrCallback(IntPtr pathname, short attributes)
        {
            OnCallBack();

            var pathNameArg = Marshal.PtrToStringBSTR(pathname);
            TrackUsage("pathname", pathNameArg, Tokens.String);
            TrackUsage("attributes", attributes, Tokens.Integer);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<SetAttrDelegate>(NativeFunctionAddress);
                nativeCall(pathname, attributes);
            }
        }
    }
}
