using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class SaveSetting : StubBase
    {
        private readonly IntPtr origAddr;
        public SaveSetting()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcSaveSetting");

            var hook = InjectDelegate(new SaveSettingDelegate(SaveSettingCallback), processAddress);
            origAddr = hook.HookBypassAddress;
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void SaveSettingDelegate(IntPtr appName, IntPtr section, IntPtr key, IntPtr setting);

        public void SaveSettingCallback(IntPtr appName, IntPtr section, IntPtr key, IntPtr setting)
        {
            OnCallBack(true);

            var appNameArg = Marshal.PtrToStringBSTR(appName);
            var sectionArg = Marshal.PtrToStringBSTR(section);
            var keyArg = Marshal.PtrToStringBSTR(key);
            var settingArg = Marshal.PtrToStringBSTR(setting);

            TrackUsage("appname", appNameArg, Tokens.String);
            TrackUsage("section", sectionArg, Tokens.String);
            TrackUsage("key", keyArg, Tokens.String);
            TrackUsage("setting", settingArg, Tokens.String);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<SaveSettingDelegate>(origAddr);
                nativeCall(appName, section, key, setting);
            }
        }
    }
}
