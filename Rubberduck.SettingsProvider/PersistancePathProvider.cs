using System;
using System.IO;

namespace Rubberduck.SettingsProvider
{
    public class PersistancePathProvider : IPersistancePathProvider
    {
        private static readonly Lazy<IPersistancePathProvider> LazyInstance;

        static PersistancePathProvider()
        {
            LazyInstance = new Lazy<IPersistancePathProvider>(() => new PersistancePathProvider());
        }

        // Disallow instancing of the class except via static method
        private PersistancePathProvider() { }

        public string DataRootPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Rubberduck");

        public string DataFolderPath(string folderName)
        {
            return Path.Combine(DataRootPath, folderName);
        }

        public static IPersistancePathProvider Instance => LazyInstance.Value;
    }
}
