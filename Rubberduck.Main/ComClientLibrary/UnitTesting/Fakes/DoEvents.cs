using Rubberduck.UnitTesting.ComClientHelpers;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class DoEvents : FakeBase
    {
        public DoEvents()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcDoEvents");

            InjectDelegate(new DoEventsDelegate(DoEventsCallback), processAddress);
        }

        private readonly ValueTypeConverter<int> _converter = new ValueTypeConverter<int>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((int)_converter.Value, invocation);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I4)]
        private delegate int DoEventsDelegate();

        public int DoEventsCallback()
        {
            OnCallBack(true);

            if (PassThrough)
            {
                return VbeProvider.VbeNativeApi.DoEvents();
            }
            return (int)(ReturnValue ?? 0);
        }
    }


}
