using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class MkDir : StubBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcMakeDir");

        public MkDir()
        {
            InjectDelegate(new MkDirDelegate(MkDirCallback), ProcessAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void MkDirDelegate(IntPtr path);

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcMakeDir(IntPtr path);

        public void MkDirCallback(IntPtr path)
        {
            OnCallBack(true);

            var pathArg = Marshal.PtrToStringBSTR(path);

            TrackUsage("path", pathArg, Tokens.String);
            if (PassThrough)
            {
                rtcMakeDir(path);
            }
        }
    }
}
