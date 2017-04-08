using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class DoEvents : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcDoEvents");

        public DoEvents()
        {
            InjectDelegate(new DoEventsDelegate(DoEventsCallback), ProcessAddress);
        }

        private readonly ValueTypeConverter<int> _converter = new ValueTypeConverter<int>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((int)_converter.Value, invocation);
        }

        [DllImport(TargetLibrary, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I4)]
        private static extern int rtcDoEvents();

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I4)]
        private delegate int DoEventsDelegate();

        public int DoEventsCallback()
        {
            OnCallBack(true);

            if (PassThrough)
            {
                return rtcDoEvents();
            }
            return (int)(ReturnValue ?? 0);
        }
    }


}
