using System.Collections.Generic;

namespace Rubberduck.SettingsProvider
{
    public interface IPersistable<T> where T : new()
    {
        void Persist(string path, IEnumerable<T> items);
        IEnumerable<T> Load(string path);
    }
}