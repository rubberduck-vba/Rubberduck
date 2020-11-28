using System;
using Path = System.IO.Path;

namespace Rubberduck.SettingsProvider
{
    internal class PersistencePathProvider : IPersistencePathProvider
    {
        private static readonly Lazy<IPersistencePathProvider> LazyInstance;

        static PersistencePathProvider()
        {
            LazyInstance = new Lazy<IPersistencePathProvider>(() => new PersistencePathProvider());
        }

        // Disallow instancing of the class except via static method
        private PersistencePathProvider() { }

        public string DataRootPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Rubberduck");

        public string DataFolderPath(string folderName)
        {
            return Path.Combine(DataRootPath, folderName);
        }

        public static IPersistencePathProvider Instance => LazyInstance.Value;
    }
}
