using System;
using System.Runtime.InteropServices;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class CurDir : FakeBase
    {
        private static readonly IntPtr ProcessAddressVariant = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcCurrentDir");
        private static readonly IntPtr ProcessAddressString = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcCurrentDirBstr");

        public CurDir()
        {
            InjectDelegate(new CurDirStringDelegate(CurDirStringCallback), ProcessAddressString);
            InjectDelegate(new CurDirVariantDelegate(CurDirVariantCallback), ProcessAddressVariant);
        }

        public override bool PassThrough
        {
            get { return false; }
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(AssertMessages.Assert_InvalidFakePassThrough, "CurDir"));
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.BStr)]
        private delegate string CurDirStringDelegate(IntPtr drive);

        public string CurDirStringCallback(IntPtr drive)
        {
            TrackInvocation(drive);
            return ReturnValue?.ToString() ?? string.Empty;
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void CurDirVariantDelegate(IntPtr retVal, IntPtr drive);

        public void CurDirVariantCallback(IntPtr retVal, IntPtr drive)
        {
            TrackInvocation(drive);
            Marshal.GetNativeVariantForObject(ReturnValue?.ToString() ?? string.Empty, retVal);
        }

        private void TrackInvocation(IntPtr drive)
        {
            OnCallBack();

            TrackUsage("drive", drive);
        }
    }
}
