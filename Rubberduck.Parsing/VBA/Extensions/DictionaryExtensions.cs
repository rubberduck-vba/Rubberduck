using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA.Extensions
{
    public static class DictionaryExtensions
    {
        public static IDictionary<TKey, TValue>ToDictionary<TKey, TValue>(this IDictionary<TKey, TValue> source)
        {
            return new Dictionary<TKey, TValue>(source);
        }
    }
}