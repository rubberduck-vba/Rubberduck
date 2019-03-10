using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class RmDir : StubBase
    {
        public RmDir()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcRemoveDir");

            InjectDelegate(new RmDirDelegate(RmDirCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void RmDirDelegate(IntPtr path);

        public void RmDirCallback(IntPtr path)
        {
            OnCallBack(true);

            var pathArg = Marshal.PtrToStringBSTR(path);

            TrackUsage("path", pathArg, Tokens.String);
            if (PassThrough)
            {
                VbeProvider.VbeNativeApi.RemoveDir(path);
            }
        }
    }
}
