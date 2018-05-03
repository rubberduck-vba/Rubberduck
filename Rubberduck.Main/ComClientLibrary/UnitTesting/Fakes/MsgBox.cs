using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class MsgBox : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcMsgBox");

        public MsgBox()
        {
            InjectDelegate(new MessageBoxDelegate(MsgBoxCallback), ProcessAddress);
        }

        public override bool PassThrough
        {
            get { return false; }
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(RubberduckUI.Assert_InvalidFakePassThrough, "MsgBox"));
            }
        }

        private readonly ValueTypeConverter<int> _converter = new ValueTypeConverter<int>();
        public override void Returns(object value, int invocation = FakesProvider.AllInvocations)
        {
            _converter.Value = value;
            base.Returns((int)_converter.Value, invocation);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate int MessageBoxDelegate(IntPtr prompt, int buttons, IntPtr title, IntPtr helpfile, IntPtr context);

        public int MsgBoxCallback(IntPtr prompt, int buttons, IntPtr title, IntPtr helpfile, IntPtr context)
        {
            OnCallBack();

            TrackUsage("prompt", prompt);
            TrackUsage("buttons", buttons, Tokens.Long);
            TrackUsage("title", title);
            TrackUsage("helpfile", helpfile);
            TrackUsage("context", context);

            return (int)(ReturnValue ?? 0);
        }    
    }
}
