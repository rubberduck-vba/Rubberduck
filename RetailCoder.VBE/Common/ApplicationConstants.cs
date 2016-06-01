using System;
using System.IO;

namespace Rubberduck.Common
{
    public static class ApplicationConstants
    {
        public static readonly string LOG_FOLDER_PATH = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck", "Logs");
    }
}
