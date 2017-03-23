using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class RmDir : StubBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcRemoveDir");

        public RmDir()
        {
            InjectDelegate(new RmDirDelegate(RmDirCallback), ProcessAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void RmDirDelegate(IntPtr path);

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern void rtcRemoveDir(IntPtr path);

        public void RmDirCallback(IntPtr path)
        {
            OnCallBack(true);

            var pathArg = Marshal.PtrToStringBSTR(path);

            TrackUsage("path", pathArg, Tokens.String);
            if (PassThrough)
            {
                rtcRemoveDir(path);
            }
        }
    }
}
