using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting
{
    internal class SendKeys : StubBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcSendKeys");

        public SendKeys()
        {
            InjectDelegate(new SendKeysDelegate(SendKeysCallback), ProcessAddress);
        }

        public override bool PassThrough
        {
            get { return false; }
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(RubberduckUI.Assert_InvalidFakePassThrough, "SendKeys"));
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void SendKeysDelegate(IntPtr String, bool wait);

        public void SendKeysCallback(IntPtr String, bool wait)
        {
            OnCallBack(true);

            var stringArg = Marshal.PtrToStringBSTR(String);

            TrackUsage("string", stringArg, Tokens.String);
            TrackUsage("wait", wait, Tokens.Boolean);
        }
    }
}
