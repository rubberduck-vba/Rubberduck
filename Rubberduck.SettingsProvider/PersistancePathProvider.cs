using System;
using System.IO;

namespace Rubberduck.SettingsProvider
{
    public class PersistancePathProvider : IPersistancePathProvider
    {
        public string DataRootPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Rubberduck");

        public string DataFolderPath(string folderName)
        {
            return Path.Combine(DataRootPath, folderName);
        }
    }
}
