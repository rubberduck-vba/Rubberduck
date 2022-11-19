using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class FileLen : FakeBase
    {
        public FileLen()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcFileLen");

            InjectDelegate(new FileLenDelegate(FileLenCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I4)]
        private delegate int FileLenDelegate(IntPtr pathname);

        public int FileLenCallback(IntPtr pathname)
        {
            OnCallBack();

            var pathNameArg = Marshal.PtrToStringBSTR(pathname);
            TrackUsage("pathname", pathNameArg, Tokens.String);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<FileLenDelegate>(NativeFunctionAddress);
                return nativeCall(pathname);
            }

            return Convert.ToInt32(ReturnValue ?? 0);
        }    
    }
}
