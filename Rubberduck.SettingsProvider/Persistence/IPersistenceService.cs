namespace Rubberduck.SettingsProvider
{
    public interface IPersistenceService<T> where T : new()
    {
        void Save(T settings, string nonDefaultFilePath = null);
        T Load(T settings, string nonDefaultFilePath = null);
    }

    public interface IFilePersistenceService<T> : IPersistenceService<T> where T : new()
    {
        string FilePath { get; set; }
    }   
}
