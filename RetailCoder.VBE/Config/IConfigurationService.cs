namespace Rubberduck.Config
{
    public interface IConfigurationService<T>
    {
        T LoadConfiguration();
        void SaveConfiguration(T toSerialize);
    }
}
