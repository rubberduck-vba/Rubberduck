namespace Rubberduck.SettingsProvider
{
    public interface IPersistencePathProvider
    {
        string DataRootPath { get; }
        string DataFolderPath(string folderName);
    }
}
