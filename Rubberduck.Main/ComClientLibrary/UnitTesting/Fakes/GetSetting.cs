using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class GetSetting : FakeBase
    {
        public GetSetting()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcGetSetting");

            InjectDelegate(new GetSettingDelegate(GetSettingCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.BStr)]
        private delegate string GetSettingDelegate(IntPtr appName, IntPtr section, IntPtr key, ComVariant.Variant defaultVal);

        public string GetSettingCallback(IntPtr appName, IntPtr section, IntPtr key, ComVariant.Variant defaultVal)
        {
            OnCallBack();

            var appNameArg = Marshal.PtrToStringBSTR(appName);
            var sectionArg = Marshal.PtrToStringBSTR(section);
            var keyArg = Marshal.PtrToStringBSTR(key);
            var defaultArg = ((VarEnum)defaultVal.vt == VarEnum.VT_BSTR) ? Marshal.PtrToStringBSTR((IntPtr)defaultVal.data01) : "";
            TrackUsage("appname", appNameArg, Tokens.String);
            TrackUsage("section", sectionArg, Tokens.String);
            TrackUsage("key", keyArg, Tokens.String);
            TrackUsage("default", defaultArg, Tokens.String); // Can't name argument as per VBA function description but keep for tracking
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<GetSettingDelegate>(NativeFunctionAddress);
                return nativeCall(appName, section, key, defaultVal);
            }

            return ReturnValue?.ToString() ?? string.Empty;
        }    
    }
}
