using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Timer : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcGetTimer");

        public Timer()
        {
            InjectDelegate(new TimerDelegate(TimerCallback), ProcessAddress);
        }

        private readonly ValueTypeConverter<float> _converter = new ValueTypeConverter<float>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((float)_converter.Value, invocation);
        }

        [DllImport(TargetLibrary, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.R4)]
        private static extern float rtcGetTimer();

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.R4)]
        private delegate float TimerDelegate();

        public float TimerCallback()
        {
            OnCallBack(true);

            if (PassThrough)
            {
                return rtcGetTimer();
            }
            return Convert.ToSingle(ReturnValue ?? 0);
        } 
    }
}
