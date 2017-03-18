using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.UnitTesting
{
    internal class MsgBox : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcMsgBox");

        public MsgBox() : base (ProcessAddress)
        {
            InjectDelegate(new MessageBoxDelegate(MsgBoxCallback));
        }

        private readonly ValueTypeConverter<int> _converter = new ValueTypeConverter<int>();
        public override void Returns(object value)
        {
            _converter.Value = value;
            base.Returns(value);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate int MessageBoxDelegate(IntPtr prompt, int buttons, IntPtr title, IntPtr helpfile, IntPtr context);

        public int MsgBoxCallback(IntPtr prompt, int buttons, IntPtr title, IntPtr helpfile, IntPtr context)
        {
            InvocationCount++;

            TrackUsage("prompt", new ComVariant(prompt));
            TrackUsage("buttons", buttons);
            TrackUsage("title", new ComVariant(title));
            TrackUsage("helpfile", new ComVariant(helpfile));
            TrackUsage("context", new ComVariant(context));

            if (Throws)
            {
                AssertHandler.RaiseVbaError(ErrorNumber, ErrorDescription);
                
            }
            return (int)_converter.Value;
        }    
    }
}
