using Rubberduck.UnitTesting.ComClientHelpers;
using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Timer : FakeBase
    {
        public Timer()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcGetTimer");

            InjectDelegate(new TimerDelegate(TimerCallback), processAddress);
        }

        private readonly ValueTypeConverter<float> _converter = new ValueTypeConverter<float>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((float)_converter.Value, invocation);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.R4)]
        private delegate float TimerDelegate();

        public float TimerCallback()
        {
            OnCallBack(true);

            if (PassThrough)
            {
                return VbeProvider.VbeNativeApi.GetTimer();
            }
            return Convert.ToSingle(ReturnValue ?? 0);
        } 
    }
}
