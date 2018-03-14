using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Shell : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcShell");

        public Shell()
        {
            InjectDelegate(new ShellDelegate(ShellCallback), ProcessAddress);
        }

        private readonly ValueTypeConverter<double> _converter = new ValueTypeConverter<double>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((double)_converter.Value, invocation);
        }

        [DllImport(TargetLibrary, SetLastError = true)]
        private static extern double rtcShell(IntPtr pathname, short windowstyle);
        
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
                return Convert.ToDouble(rtcShell(pathname, windowstyle));
            }
            return Convert.ToDouble(ReturnValue ?? 0);
        }
    }
}
