using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UnitTesting.ComClientHelpers;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Shell : FakeBase
    {
        public Shell()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcShell");

            InjectDelegate(new ShellDelegate(ShellCallback), processAddress);
        }

        private readonly ValueTypeConverter<double> _converter = new ValueTypeConverter<double>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((double)_converter.Value, invocation);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.R8)]
        private delegate double ShellDelegate(IntPtr pathname, short windowstyle);

        public double ShellCallback(IntPtr pathname, short windowstyle)
        {
            OnCallBack(true);

            var path = Marshal.PtrToStringBSTR(pathname);

            TrackUsage("pathname", pathname);
            TrackUsage("windowstyle", windowstyle, Tokens.Integer);

            if (PassThrough)
            {
                return Convert.ToDouble(VbeProvider.VbeNativeApi.Shell(pathname, windowstyle));
            }
            return Convert.ToDouble(ReturnValue ?? 0);
        }
    }
}
