using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.UnitTesting.ComClientHelpers;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class MsgBox : FakeBase
    {
        public MsgBox()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcMsgBox");

            InjectDelegate(new MessageBoxDelegate(MsgBoxCallback), processAddress);
        }

        public override bool PassThrough
        {
            get { return false; }
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(AssertMessages.Assert_InvalidFakePassThrough, "MsgBox"));
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
