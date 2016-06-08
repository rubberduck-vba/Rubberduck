using System.Diagnostics;

namespace Rubberduck.Common
{
    public sealed class WindowsOperatingSystem : IOperatingSystem
    {
        public void ShowFolder(string folderPath)
        {
            Process.Start(folderPath);
        }
    }
}
