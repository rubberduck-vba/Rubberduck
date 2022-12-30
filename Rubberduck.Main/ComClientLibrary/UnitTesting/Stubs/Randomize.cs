using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    internal class Randomize : StubBase
    {
        public Randomize()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcRandomize");

            InjectDelegate(new RandomizeDelegate(RandomizeCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void RandomizeDelegate(IntPtr number);

        public void RandomizeCallback(IntPtr number)
        {
            OnCallBack();

            TrackUsage("number", number);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<RandomizeDelegate>(NativeFunctionAddress);
                nativeCall(number);
            }
        }
    }
}
