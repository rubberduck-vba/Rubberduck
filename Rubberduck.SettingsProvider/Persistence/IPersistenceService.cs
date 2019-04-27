namespace Rubberduck.SettingsProvider
{
    /// <summary>
    /// An interface exposing persistent storage for a serializable type.
    /// This is intended to be used by Settings only.
    /// 
    /// When the optional parameters are given, they override the storage location of the instance state for that one invocation.
    /// Implementations of this interface are expected to <b>not</b> provide any caching.
    /// 
    /// Implementations can choose to disregard the optional argument being
    /// </summary>
    /// <typeparam name="T">The Type of serializable object to store and retrieve persistently</typeparam>
    public interface IPersistenceService<T> where T : new()
    {
        void Save(T settings, string path = null);
        T Load(string path = null);
    }
}
