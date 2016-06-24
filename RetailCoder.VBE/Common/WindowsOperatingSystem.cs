using System.Diagnostics;
using System.IO;

namespace Rubberduck.Common
{
    public sealed class WindowsOperatingSystem : IOperatingSystem
    {
        public void ShowFolder(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            Process.Start(folderPath);
        }
    }
}
