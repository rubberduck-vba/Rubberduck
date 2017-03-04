namespace Rubberduck.SettingsProvider
{
    public interface IPersistanceService<T> where T : new()
    {
        void Save(T settings);
        T Load(T settings);
    }

    public interface IFilePersistanceService<T> : IPersistanceService<T> where T : new()
    {
        string FilePath { get; set; }
    }   
}
