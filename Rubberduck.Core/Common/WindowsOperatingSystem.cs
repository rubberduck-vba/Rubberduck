using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.CompilerServices;
using NLog;

namespace Rubberduck.Common
{
    public sealed class WindowsOperatingSystem : IOperatingSystem
    {
        public static readonly ILogger _Logger = LogManager.GetCurrentClassLogger(); 

        public void ShowFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            using (Process.Start(folderPath))
            {
            }
        }


        public WindowsVersion? GetOSVersion()
        {
            try
            {
                var wmiEnum = new ManagementObjectSearcher("root\\CIMV2", "SELECT Version FROM  Win32_OperatingSystem")
                    .Get().GetEnumerator();
                wmiEnum.MoveNext();
                var versionString = wmiEnum.Current.Properties["Version"].Value as string;

                var versionElements = versionString?.Split('.');

                if (versionElements?.Length >= 3 &&
                    int.TryParse(versionElements[0], out var major) &&
                    int.TryParse(versionElements[1], out var minor) &&
                    int.TryParse(versionElements[2], out var build))
                {
                    return new WindowsVersion(major, minor, build);
                }                
            }
            catch (Exception exception)
            {
                _Logger.Warn(exception, "Unable to determine OS Version");
                return null;
            }
            _Logger.Warn("Unable to determine OS Version");
            return null;
        }
    }
}

