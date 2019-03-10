using System;
using System.Runtime.InteropServices;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class CurDir : FakeBase
    {
        public CurDir()
        {
            var processAddressVariant = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcCurrentDir");
            var processAddressString = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcCurrentDirBstr");

            InjectDelegate(new CurDirStringDelegate(CurDirStringCallback), processAddressString);
            InjectDelegate(new CurDirVariantDelegate(CurDirVariantCallback), processAddressVariant);
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
