using System;
using System.Runtime.InteropServices;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting.Fakes
{
    // TODO: This is currently broken.  The runtime throws a "bad dll calling convention error when it returns", which leads me to
    // believe that it is trying to cast the return value and it isn't marshalling correctly.
    internal class CurDir : FakeBase
    {
        private static readonly IntPtr ProcessAddressString = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcCurrentDir");
        private static readonly IntPtr ProcessAddressVariant = EasyHook.LocalHook.GetProcAddress(TargetLibrary, "rtcCurrentDirBstr");

        public CurDir()
        {
            InjectDelegate(new CurDirStringDelegate(CurDirStringCallback), ProcessAddressString);
            InjectDelegate(new CurDirVariantDelegate(CurDirStringCallback), ProcessAddressVariant);
        }

        public override bool PassThrough
        {
            get { return false; }
            // ReSharper disable once ValueParameterNotUsed
            set
            {
                Verifier.SuppressAsserts();
                AssertHandler.OnAssertInconclusive(string.Format(RubberduckUI.Assert_InvalidFakePassThrough, "CurDir"));
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate string CurDirStringDelegate(IntPtr drive);

        public string CurDirStringCallback(IntPtr drive)
        {
            TrackInvocation(drive);
            return ReturnValue?.ToString() ?? string.Empty;
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate string CurDirVariantDelegate(IntPtr drive);

        public object CurDirVariantCallback(IntPtr drive)
        {
            TrackInvocation(drive);
            return ReturnValue?.ToString() ?? string.Empty;
        }

        private void TrackInvocation(IntPtr drive)
        {
            OnCallBack();

            TrackUsage("drive", drive);
        }
    }
}
