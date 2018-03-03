using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class ChDir : StubBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcChangeDir");

        public ChDir()
        {
            InjectDelegate(new ChDirDelegate(ChDirCallback), ProcessAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void ChDirDelegate(IntPtr path);

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcChangeDir(IntPtr path);

        public void ChDirCallback(IntPtr path)
        {
            OnCallBack(true);

            var pathArg = Marshal.PtrToStringBSTR(path);

            TrackUsage("path", pathArg, Tokens.String);
            if (PassThrough)
            {
                rtcChangeDir(path);
            }
        }
    }
}
