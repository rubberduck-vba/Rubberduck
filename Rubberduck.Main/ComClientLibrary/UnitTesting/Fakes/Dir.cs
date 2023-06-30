using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Dir : FakeBase
    {
        public Dir()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcDir");

            InjectDelegate(new DirDelegate(DirCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.BStr)]
        private delegate string DirDelegate(IntPtr pathname, short attributes);

        public string DirCallback(IntPtr pathname, short attributes)
        {
            OnCallBack();

            TrackUsage("pathname", pathname);
            TrackUsage("attributes", attributes, Tokens.Integer);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<DirDelegate>(NativeFunctionAddress);
                return nativeCall(pathname, attributes);
            }

            return ReturnValue?.ToString() ?? string.Empty;
        }    
    }
}
