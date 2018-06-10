using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Now : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcGetPresentDate");

        public Now()
        {
            InjectDelegate(new NowDelegate(NowCallback), ProcessAddress);
        }

        private readonly ValueTypeConverter<float> _converter = new ValueTypeConverter<float>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((float)_converter.Value, invocation);
        }

        [DllImport(TargetLibrary, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.R4)]
        private static extern void rtcGetPresentDate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void NowDelegate(IntPtr retVal);

        public void NowCallback(IntPtr retVal)
        {
            OnCallBack(true);

            if (PassThrough)
            {
                //TODO - get passthrough working with byref call
                Marshal.GetNativeVariantForObject(0, retVal); // rtcGetPresentDate();
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
