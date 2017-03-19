using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.UnitTesting
{
    internal class InputBox : FakeBase
    {
        private static readonly IntPtr ProcessAddress = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcInputBox");

        public InputBox() : base (ProcessAddress)
        {
            InjectDelegate(new InputBoxDelegate(InputBoxCallback));
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.BStr)]
        private delegate string InputBoxDelegate(IntPtr prompt, IntPtr title, IntPtr Default, IntPtr xpos, IntPtr ypos, IntPtr helpfile, IntPtr context);

        public string InputBoxCallback(IntPtr prompt, IntPtr title, IntPtr Default, IntPtr xpos, IntPtr ypos, IntPtr helpfile, IntPtr context)
        {
            OnCallBack();

            TrackUsage("prompt", new ComVariant(prompt));
            TrackUsage("title", new ComVariant(title));
            TrackUsage("default", new ComVariant(Default));
            TrackUsage("xpos", new ComVariant(xpos));
            TrackUsage("ypos", new ComVariant(ypos));
            TrackUsage("helpfile", new ComVariant(helpfile));
            TrackUsage("context", new ComVariant(context));

            return ReturnValue?.ToString() ?? string.Empty;
        }
    }
}
