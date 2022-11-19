using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class FileDateTime : FakeBase
    {
        public FileDateTime()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcFileDateTime");

            InjectDelegate(new FileDateTimeDelegate(FileDateTimeCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void FileDateTimeDelegate(IntPtr retVal, IntPtr pathname);

        public void FileDateTimeCallback(IntPtr retVal, IntPtr pathname)
        {
            OnCallBack();

            var pathNameArg = Marshal.PtrToStringBSTR(pathname);
            TrackUsage("pathname", pathNameArg, Tokens.String);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<FileDateTimeDelegate>(NativeFunctionAddress);
                nativeCall(retVal, pathname);
            }

            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }    
    }
}
