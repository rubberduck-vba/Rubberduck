using System;
using Microsoft.Win32;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBERuntime
{
    public class VBESettings : IVBESettings
    {
        private const string Vbe7SettingPath = @"HKEY_CURRENT_USER\Software\Microsoft\VBA\7.0\Common";
        private const string Vbe6SettingPath = @"HKEY_CURRENT_USER\Software\Microsoft\VBA\6.0\Common";

        public enum DllVersion
        {
            Unknown,
            Vbe6,
            Vbe7
        }

        private readonly string _registryRootPath;

        public VBESettings(IVBE vbe)
        {
            try
            {
                switch (Convert.ToInt32(decimal.Parse(vbe.Version)))
                {
                    case 6:
                        Version = DllVersion.Vbe6;
                        _registryRootPath = Vbe6SettingPath;
                        break;
                    case 7:
                        Version = DllVersion.Vbe7;
                        _registryRootPath = Vbe7SettingPath;
                        break;
                    default:
                        Version = DllVersion.Unknown;
                        break;
                }
            }
            catch
            {
                Version = DllVersion.Unknown;
                _registryRootPath = null;
            }
        }

        public DllVersion Version { get; }

        public bool CompileOnDemand
        {
            get => DWordToBooleanConverter(nameof(CompileOnDemand));
            set => BooleanToDWordConverter(nameof(CompileOnDemand), value);
        }

        public bool BackGroundCompile
        {
            get => DWordToBooleanConverter(nameof(BackGroundCompile));
            set => BooleanToDWordConverter(nameof(BackGroundCompile), value);
        }

        private const int DWordTrueValue = 1;
        private const int DWordFalseValue = 0;

        private bool DWordToBooleanConverter(string keyName)
        {
            var result = Registry.GetValue(_registryRootPath, keyName, DWordFalseValue) as int?;
            return Convert.ToBoolean(result ?? DWordFalseValue);
        }

        private void BooleanToDWordConverter(string keyName, bool value)
        {
            Registry.SetValue(_registryRootPath, keyName, value ? DWordTrueValue : DWordFalseValue, RegistryValueKind.DWord);
        }
    }
}
