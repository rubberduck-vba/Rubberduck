using System;
using System.Collections.Concurrent;

namespace Rubberduck.Parsing.ComReflection.TypeLibReflection
{
    internal interface ILibraryTypeCache
    {
        string Key { get; }
        bool TryGetType(string progId, out Type type);
        bool AddType(string progId, Type type);
        Type GetOrAdd(string progId, Type type);
        bool Remove(string progId);
    }

    internal sealed class LibraryTypeCache : ILibraryTypeCache
    {
        private readonly ConcurrentDictionary<string, Type> _cache;

        public LibraryTypeCache(string key)
        {
            Key = key;
            _cache = new ConcurrentDictionary<string, Type>();
        }

        public string Key { get; }

        public bool TryGetType(string progId, out Type type)
        {
            return _cache.TryGetValue(progId.ToLowerInvariant(), out type);
        }

        public bool AddType(string progId, Type type)
        {
            if (_cache.ContainsKey(progId.ToLowerInvariant()))
            {
                return false;
            }

            _cache.AddOrUpdate(progId.ToLowerInvariant(), p => type, (p, t) => type);
            return true;
        }

        public Type GetOrAdd(string progId, Type type)
        {
            return _cache.GetOrAdd(progId.ToLowerInvariant(), s => type);
        }

        public bool Remove(string progId)
        {
            return _cache.TryRemove(progId, out _);
        }
    }
}
