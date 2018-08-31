using System;
using System.Runtime.InteropServices;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class Environ : FakeBase
    {
        private static readonly IntPtr ProcessAddressString = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcEnvironBstr");
        private static readonly IntPtr ProcessAddressVariant = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcEnvironVar");

        public Environ()
        {
            InjectDelegate(new EnvironStringDelegate(EnvironStringCallback), ProcessAddressString);
            InjectDelegate(new EnvironVariantDelegate(EnvironVariantCallback), ProcessAddressVariant);
        }

        public override bool PassThrough
        {
            get { return false; }
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(AssertMessages.Assert_InvalidFakePassThrough, "Environ"));
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.BStr)]
        private delegate string EnvironStringDelegate(IntPtr envstring);

        public string EnvironStringCallback(IntPtr envstring)
        {
            TrackInvocation(envstring);
            return ReturnValue?.ToString() ?? string.Empty;
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void EnvironVariantDelegate(IntPtr retVal, IntPtr envstring);

        public void EnvironVariantCallback(IntPtr retVal, IntPtr envstring)
        {
            TrackInvocation(envstring);
            Marshal.GetNativeVariantForObject(ReturnValue?.ToString() ?? string.Empty, retVal);
        }

        private void TrackInvocation(IntPtr envstring)
        {
            OnCallBack();

            TrackUsage("envstring", envstring);
        }
    }
}
