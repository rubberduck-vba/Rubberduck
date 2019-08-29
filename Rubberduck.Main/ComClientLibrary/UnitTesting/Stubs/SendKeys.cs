using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UnitTesting
{
    internal class SendKeys : StubBase
    {
        public SendKeys()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcSendKeys");

            InjectDelegate(new SendKeysDelegate(SendKeysCallback), processAddress);
        }

        public override bool PassThrough
        {
            get => false;
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(AssertMessages.Assert_InvalidFakePassThrough, "SendKeys"));
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
