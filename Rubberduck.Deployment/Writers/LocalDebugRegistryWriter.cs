using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.Win32;
using Rubberduck.Deployment.Structs;

namespace Rubberduck.Deployment.Writers
{
    public class LocalDebugRegistryWriter : IRegistryWriter
    {
        public string Write(IOrderedEnumerable<RegistryEntry> entries)
        {
            //System.Diagnostics.Debugger.Launch(); //uncomment if need to debug
            var sb = new StringBuilder("Windows Registry Editor Version 5.00" + Environment.NewLine + Environment.NewLine);
            var distinctKeys = new List<string>();
            
            foreach (var entry in entries)
            {
                //Guard clause to prevent registry write to wrong places
                if (!entry.Key.StartsWith("Software\\Classes") &&
                    !entry.Key.StartsWith("Software\\Classes\\CLSID\\") &&
                    !entry.Key.StartsWith("Software\\Classes\\Interface\\") &&
                    !entry.Key.StartsWith("Software\\Classes\\TypeLib\\") &&
                    !entry.Key.StartsWith("Software\\Classes\\Record\\") &&
                    !entry.Key.StartsWith("Software\\Classes\\Rubberduck.")
                )
                {
                    throw new InvalidOperationException("Unexpected registry entry: " + entry.Key);
                }

                if (Environment.Is64BitOperatingSystem)
                {
                    MakeRegistryEntries(entry, RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32));
                    MakeRegistryEntries(entry, RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64));
                }
                else 
                {
                    MakeRegistryEntries(entry, Registry.CurrentUser);
                }

                if (!distinctKeys.Contains(entry.Key))
                {
                    distinctKeys.Add(entry.Key);
                }
            }

            foreach (var key in distinctKeys)
            {
                //we need a break each entry, so 2 newline is wanted (AppendLine adds one, and we add another)
                sb.AppendLine("[-HKEY_CURRENT_USER\\" + key + "]" + Environment.NewLine);
            }

            return sb.ToString();
        }

        private void MakeRegistryEntries(RegistryEntry entry, RegistryKey hKey) 
        {
            var key = hKey.CreateSubKey(entry.Key);

            var value = ReplacePlaceholder(entry.Value, entry.Bitness);

            if (!(string.IsNullOrWhiteSpace(entry.Name) && string.IsNullOrWhiteSpace(value)))
            {
                key.SetValue(entry.Name, value);
            }
        }

        //Cache the string so we call the AssemblyDirectory only once
        private string _currentPath;
        private string ReplacePlaceholder(string value, Bitness bitness)
        {
            if (_currentPath == null)
            {
                _currentPath = AssemblyDirectory;
            }

            switch (value)
            {
                case PlaceHolders.InstallPath:
                    return _currentPath;
                case PlaceHolders.DllPath:
                    return Path.Combine( _currentPath, "Rubberduck.dll");
                case PlaceHolders.TlbPath:
                    return Path.Combine(_currentPath,  bitness == Bitness.Is64Bit ? "Rubberduck.x64.tlb" : "Rubberduck.x32.tlb");
                default:
                    return value;
            }
        }

        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }
    }
}
