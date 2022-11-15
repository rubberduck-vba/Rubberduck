using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.UnitTesting.Fakes
{
    internal class GetSetting : FakeBase
    {
        private readonly IntPtr origAddr;
        public GetSetting()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcGetSetting");

            var hook = InjectDelegate(new GetSettingDelegate(GetSettingCallback), processAddress);
            origAddr = hook.HookBypassAddress;
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.BStr)]
        private delegate string GetSettingDelegate(IntPtr appName, IntPtr section, IntPtr key, ComVariant.Variant defaultVal);

        public string GetSettingCallback(IntPtr appName, IntPtr section, IntPtr key, ComVariant.Variant defaultVal)
        {
            OnCallBack(true);

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
                var nativeCall = Marshal.GetDelegateForFunctionPointer<GetSettingDelegate>(origAddr);
                return nativeCall(appName, section, key, defaultVal);
            }

            return ReturnValue?.ToString() ?? string.Empty;
        }    
    }
}
