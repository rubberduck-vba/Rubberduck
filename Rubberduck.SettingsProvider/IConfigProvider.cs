namespace Rubberduck.SettingsProvider
{
    public interface IConfigProvider<T>
    {
        T Create();
        T CreateDefaults();

        void Save(T settings);
    }
}