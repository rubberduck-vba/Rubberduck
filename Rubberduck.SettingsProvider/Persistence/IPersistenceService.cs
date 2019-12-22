namespace Rubberduck.SettingsProvider
{
    /// <summary>
    /// An interface exposing persistent storage for a serializable type.
    /// This is intended to be used by Settings only.
    /// </summary>
    /// <remarks>
    /// When the optional parameters are given, they override the storage location of the instance state for that one invocation.
    /// Implementations of this interface are expected to <b>not</b> provide any caching.
    /// The following property must hold for a given instance of IPersistenceService:
    /// <code>
    /// T value;
    /// string arg; // can be null
    /// IPersistenceService<T> service;
    /// service.Save(value, arg);
    /// service.Load(arg).Equals(value); // must be true
    /// </code>
    /// </remarks>
    /// <typeparam name="T">The Type of serializable object to store and retrieve persistently</typeparam>
    public interface IPersistenceService<T> where T : new()
    {
        void Save(T settings, string path = null);
        T Load(string path = null);
    }
}
