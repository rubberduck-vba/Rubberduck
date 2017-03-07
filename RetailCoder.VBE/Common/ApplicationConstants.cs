using System;
using System.IO;

namespace Rubberduck.Common
{
    public static class ApplicationConstants
    {
        public static readonly string RUBBERDUCK_FOLDER_PATH = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");
        public static readonly string LOG_FOLDER_PATH = Path.Combine(RUBBERDUCK_FOLDER_PATH, "Logs");
        public static readonly string RUBBERDUCK_TEMP_PATH = Path.Combine(Path.GetTempPath(), "Rubberduck");
    }
}
