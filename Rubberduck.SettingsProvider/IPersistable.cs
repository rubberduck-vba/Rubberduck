namespace Rubberduck.SettingsProvider
{
    public interface IPersistable<T> where T : class
    {
        void Persist(string path, T tree);
        T Load(string path);
    }
}