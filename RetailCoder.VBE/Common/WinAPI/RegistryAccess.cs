﻿using Microsoft.Win32;

namespace Rubberduck.Common.WinAPI
{
    internal static class RegistryAccess
    {
        internal static RegistryKey GetDeviceKey(string device)
        {
            var split = device.Substring(4).Split('#');

            var classCode = split[0];       // ACPI (Class code)
            var subClassCode = split[1];    // PNP0303 (SubClass code)
            var protocolCode = split[2];    // 3&13c0b0c5&0 (Protocol code)

            return Registry.LocalMachine.OpenSubKey(
                $@"System\CurrentControlSet\Enum\{classCode}\{subClassCode}\{protocolCode}");
        }

        internal static string GetClassType(string classGuid)
        {
            var classGuidKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Class\" + classGuid);

            return classGuidKey != null ? (string)classGuidKey.GetValue("Class") : string.Empty;
        }
    }
}
