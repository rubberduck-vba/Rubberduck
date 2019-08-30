using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class MkDir : StubBase
    {
        public MkDir()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcMakeDir");

            InjectDelegate(new MkDirDelegate(MkDirCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void MkDirDelegate(IntPtr path);

        public void MkDirCallback(IntPtr path)
        {
            OnCallBack(true);

            var pathArg = Marshal.PtrToStringBSTR(path);

            TrackUsage("path", pathArg, Tokens.String);
            if (PassThrough)
            {
                VbeProvider.VbeNativeApi.MakeDir(path);
            }
        }
    }
}
