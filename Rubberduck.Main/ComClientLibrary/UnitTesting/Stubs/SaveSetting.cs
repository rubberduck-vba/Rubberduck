using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UnitTesting
{
    internal class SaveSetting : StubBase
    {
        public SaveSetting()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcSaveSetting");

            InjectDelegate(new SaveSettingDelegate(SaveSettingCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void SaveSettingDelegate(IntPtr appName, IntPtr section, IntPtr key, IntPtr setting);

        public void SaveSettingCallback(IntPtr appName, IntPtr section, IntPtr key, IntPtr setting)
        {
            OnCallBack();

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
                var nativeCall = Marshal.GetDelegateForFunctionPointer<SaveSettingDelegate>(NativeFunctionAddress);
                nativeCall(appName, section, key, setting);
            }
        }
    }
}
