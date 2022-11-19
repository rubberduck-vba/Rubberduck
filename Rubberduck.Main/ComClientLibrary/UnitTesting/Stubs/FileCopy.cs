using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class FileCopy : StubBase
    {
        public FileCopy()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcFileCopy");

            InjectDelegate(new FileCopyDelegate(FileCopyCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void FileCopyDelegate(IntPtr oldpathname, IntPtr newpathname);

        public void FileCopyCallback(IntPtr oldpathname, IntPtr newpathname)
        {
            OnCallBack();

            var oldpathnameArg = Marshal.PtrToStringBSTR(oldpathname);
            var newpathnameArg = Marshal.PtrToStringBSTR(newpathname);
            TrackUsage("oldpathname", oldpathnameArg, Tokens.String);
            TrackUsage("newpathname", newpathnameArg, Tokens.String);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<FileCopyDelegate>(NativeFunctionAddress);
                nativeCall(oldpathname, newpathname);
            }
        }
    }
}
