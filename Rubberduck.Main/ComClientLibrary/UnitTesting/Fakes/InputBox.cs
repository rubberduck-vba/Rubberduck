using System;
using System.Runtime.InteropServices;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class InputBox : FakeBase
    {
        public InputBox()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcInputBox");

            InjectDelegate(new InputBoxDelegate(InputBoxCallback), processAddress);
        }

        public override bool PassThrough
        {
            get { return false; }
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(AssertMessages.Assert_InvalidFakePassThrough, "InputBox"));
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.BStr)]
        private delegate string InputBoxDelegate(IntPtr prompt, IntPtr title, IntPtr Default, IntPtr xpos, IntPtr ypos, IntPtr helpfile, IntPtr context);

        public string InputBoxCallback(IntPtr prompt, IntPtr title, IntPtr Default, IntPtr xpos, IntPtr ypos, IntPtr helpfile, IntPtr context)
        {
            OnCallBack();

            TrackUsage("prompt", prompt);
            TrackUsage("title", title);
            TrackUsage("default", Default);
            TrackUsage("xpos", xpos);
            TrackUsage("ypos", ypos);
            TrackUsage("helpfile", helpfile);
            TrackUsage("context", context);

            return ReturnValue?.ToString() ?? string.Empty;
        }
    }
}
