using System;
using Microsoft.Win32;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.VBEditor.VbeRuntime.Settings
{
    public class VbeSettings : IVbeSettings
    {
        private const string Vbe7SettingPath = @"HKEY_CURRENT_USER\Software\Microsoft\VBA\7.0\Common";
        private const string Vbe6SettingPath = @"HKEY_CURRENT_USER\Software\Microsoft\VBA\6.0\Common";

        private readonly IRegistryWrapper _registry;
        private readonly string _activeRegistryRootPath;
        private readonly string[] _registryRootPaths = { Vbe7SettingPath, Vbe6SettingPath };

        public VbeSettings(IVBE vbe, IRegistryWrapper registry)
        {
            try
            {
                switch (VbeDllVersion.GetCurrentVersion(vbe))
                {
                    case DllVersion.Vbe6:
                        Version = DllVersion.Vbe6;
                        _activeRegistryRootPath = Vbe6SettingPath;
                        break;
                    case DllVersion.Vbe7:
                        Version = DllVersion.Vbe7;
                        _activeRegistryRootPath = Vbe7SettingPath;
                        break;
                    default:
                        Version = DllVersion.Unknown;
                        break;
                }
            }
            catch
            {
                Version = DllVersion.Unknown;
                _activeRegistryRootPath = null;
            }
            _registry = registry;
        }

        public DllVersion Version { get; }

        public bool CompileOnDemand
        {
            get => ReadActiveRegistryPath(nameof(CompileOnDemand));
            set => WriteAllRegistryPaths(nameof(CompileOnDemand), value);
        }

        public bool BackGroundCompile
        {
            get => ReadActiveRegistryPath(nameof(BackGroundCompile));
            set => WriteAllRegistryPaths(nameof(BackGroundCompile), value);
        }

        private bool ReadActiveRegistryPath(string keyName)
        {
            return DWordToBooleanConverter(_activeRegistryRootPath, keyName) ?? false;
        }

        private void WriteAllRegistryPaths(string keyName, bool value)
        {
            foreach (var path in _registryRootPaths)
            {
                if (DWordToBooleanConverter(path, keyName) != null)
                {
                    BooleanToDWordConverter(path, keyName, value);
                }
            }
        }

        private const int DWordTrueValue = 1;
        private const int DWordFalseValue = 0;

        private bool? DWordToBooleanConverter(string path, string keyName)
        {
            return !(_registry.GetValue(path, keyName, DWordFalseValue) is int result)
                ? (bool?)null
                : Convert.ToBoolean(result);
        }

        private void BooleanToDWordConverter(string path, string keyName, bool value)
        {
            _registry.SetValue(path, keyName, value ? DWordTrueValue : DWordFalseValue, RegistryValueKind.DWord);
        }
    }
}
