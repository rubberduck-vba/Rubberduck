namespace Rubberduck.Settings
{
    public interface IConfigProvider<T>
    {
        T Create();
        T CreateDefaults();

        void Save(T settings);
    }
}