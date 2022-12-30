using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class GetAllSettings : FakeBase
    {
        public GetAllSettings()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcGetAllSettings");

            InjectDelegate(new GetAllSettingsDelegate(GetAllSettingsCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void GetAllSettingsDelegate(IntPtr retVal, IntPtr appName, IntPtr section);

        public void GetAllSettingsCallback(IntPtr retVal, IntPtr appName, IntPtr section)
        {
            OnCallBack();

            var appNameArg = Marshal.PtrToStringBSTR(appName);
            var sectionArg = Marshal.PtrToStringBSTR(section);
            TrackUsage("appname", appNameArg, Tokens.String);
            TrackUsage("section", sectionArg, Tokens.String);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<GetAllSettingsDelegate>(NativeFunctionAddress);
                nativeCall(retVal, appName, section);
                return;
            }
            Marshal.GetNativeVariantForObject(ReturnValue ?? 0, retVal);
        }
    }
}
