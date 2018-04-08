using System;
using System.Runtime.InteropServices;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Environ : FakeBase
    {
        private static readonly IntPtr ProcessAddressString = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcEnvironBstr");
        private static readonly IntPtr ProcessAddressVariant = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcEnvironVar");

        public Environ()
        {
            InjectDelegate(new EnvironStringDelegate(EnvironStringCallback), ProcessAddressString);
            InjectDelegate(new EnvironVariantDelegate(EnvironStringCallback), ProcessAddressVariant);
        }

        public override bool PassThrough
        {
            get { return false; }
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(RubberduckUI.Assert_InvalidFakePassThrough, "Environ"));
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.BStr)]
        private delegate string EnvironStringDelegate(IntPtr envstring, IntPtr number);

        public string EnvironStringCallback(IntPtr envstring, IntPtr number)
        {
            TrackInvocation(envstring, number);              
            return ReturnValue?.ToString() ?? string.Empty;
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate string EnvironVariantDelegate(IntPtr envstring, IntPtr number);

        public object EnvironVariantCallback(IntPtr envstring, IntPtr number)
        {
            TrackInvocation(envstring, number);
            return ReturnValue?.ToString() ?? string.Empty;
        }

        private void TrackInvocation(IntPtr envstring, IntPtr number)
        {
            OnCallBack();

            TrackUsage("envstring", envstring);
            TrackUsage("number", number);
        }
    }
}
