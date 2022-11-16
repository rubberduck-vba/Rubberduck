using System;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.UnitTesting
{
    internal class DeleteSetting : StubBase
    {
        public DeleteSetting()
        {
            var processAddress = EasyHook.LocalHook.GetProcAddress(VbeProvider.VbeNativeApi.DllName, "rtcDeleteSetting");

            InjectDelegate(new DeleteSettingDelegate(DeleteSettingCallback), processAddress);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate void DeleteSettingDelegate(IntPtr appName, ComVariant.Variant section, ComVariant.Variant key);

        public void DeleteSettingCallback(IntPtr appName, ComVariant.Variant section, ComVariant.Variant key)
        {
            OnCallBack();

            var appNameArg = Marshal.PtrToStringBSTR(appName);
            var sectionArg = Marshal.PtrToStringBSTR((IntPtr)section.data01);
            var keyArg = ((VarEnum)key.vt == VarEnum.VT_BSTR) ? Marshal.PtrToStringBSTR((IntPtr)key.data01) : "";

            TrackUsage("appname", appNameArg, Tokens.String);
            TrackUsage("section", sectionArg, Tokens.String);
            TrackUsage("key", keyArg, Tokens.String);
            if (PassThrough)
            {
                var nativeCall = Marshal.GetDelegateForFunctionPointer<DeleteSettingDelegate>(NativeFunctionAddress);
                nativeCall(appName, section, key);
            }
        }
    }
}
