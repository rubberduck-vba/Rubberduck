namespace Rubberduck.SettingsProvider
{
    public interface IPersistanceService<T> where T : new()
    {
        void Save(T settings, string nonDefaultFilePath = null);
        T Load(T settings, string nonDefaultFilePath = null);
    }

    public interface IFilePersistanceService<T> : IPersistanceService<T> where T : new()
    {
        string FilePath { get; set; }
    }   
}
