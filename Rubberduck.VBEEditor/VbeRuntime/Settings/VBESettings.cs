using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.VBEditor.VbeRuntime.Settings
{
    public class VbeSettings : IVbeSettings
    {
        private static readonly List<string> VbeVersions = new List<string> { "6.0", "7.0", "7.1" };
        private const string VbeSettingPathTemplate = @"HKEY_CURRENT_USER\Software\Microsoft\VBA\{0}\Common";
        private const string Vb6SettingPath = @"HKEY_CURRENT_USER\Software\Microsoft\VBA\Microsoft Visual Basic";

        private readonly IRegistryWrapper _registry;
        private readonly string _activeRegistryRootPath;
        

        public VbeSettings(IVBE vbe, IRegistryWrapper registry)
        {
            try
            {
                Version = VbeDllVersion.GetCurrentVersion(vbe);
                switch (Version)
                {
                    case DllVersion.Vbe7:
                    case DllVersion.Vbe6:
                        _activeRegistryRootPath = string.Format(VbeSettingPathTemplate, vbe.Version.Substring(0, 3));
                        break;
                    case DllVersion.Vb98:
                        _activeRegistryRootPath = Vb6SettingPath;
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
            var paths = VbeVersions.Select(version => string.Format(VbeSettingPathTemplate, version))
                .Union(new[] {Vb6SettingPath});

            foreach (var path in paths)
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
