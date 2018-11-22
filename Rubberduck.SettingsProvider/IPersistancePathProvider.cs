namespace Rubberduck.SettingsProvider
{
    public interface IPersistancePathProvider
    {
        string DataRootPath { get; }
        string DataFolderPath(string folderName);
    }
}
