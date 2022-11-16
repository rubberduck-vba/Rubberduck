using Rubberduck.UnitTesting.ComClientHelpers;
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Rnd : FakeBase
    {
        public Rnd()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcRandomNext");

            InjectDelegate(new RndDelegate(RndCallback), processAddress);
        }

        private readonly ValueTypeConverter<float> _converter = new ValueTypeConverter<float>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((float)_converter.Value, invocation);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.R4)]
        private delegate float RndDelegate(IntPtr inVal);

        public float RndCallback(IntPtr number)
        {
            OnCallBack(true);

            TrackUsage("number", number);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<RndDelegate>(NativeFunctionAddress);
                return nativeCall(number);
            }
            return Convert.ToSingle(ReturnValue ?? 0);
        } 
    }
}
