using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.VBEditor.Extensions
{
    public static class DictionaryExtensions
    {
        public static IEnumerable<TValue> AllValues<TKey, TValue>(
            this ConcurrentDictionary<TKey, ConcurrentBag<TValue>> source)
        {
            return source.SelectMany(item => item.Value).ToList();
        }

        public static IEnumerable<TValue> AllValues<TKey, TValue>(
            this IDictionary<TKey, IList<TValue>> source)
        {
            return source.SelectMany(item => item.Value).ToList();
        }

        public static IEnumerable<TValue> AllValues<TKey, TValue>(
            this IDictionary<TKey, List<TValue>> source)
        {
            return source.SelectMany(item => item.Value);
        }

        public static ConcurrentDictionary<TKey, ConcurrentBag<TValue>> ToConcurrentDictionary<TKey, TValue>(this IEnumerable<IGrouping<TKey, TValue>> source)
        {
            return new ConcurrentDictionary<TKey, ConcurrentBag<TValue>>(source.Select(x => new KeyValuePair<TKey, ConcurrentBag<TValue>>(x.Key, new ConcurrentBag<TValue>(x))));
        }

        public static Dictionary<TKey, List<TValue>> ToDictionary<TKey, TValue>(this IEnumerable<IGrouping<TKey, TValue>> source)
        {
            return source.ToDictionary(group => group.Key, group => group.ToList());
        }

        public static IReadOnlyDictionary<TKey, IReadOnlyList<TValue>> ToReadonlyDictionary<TKey, TValue>(this IEnumerable<IGrouping<TKey, TValue>> source)
        {
            return source.ToDictionary(group => group.Key, group => (IReadOnlyList<TValue>)group.ToList());
        }

        //See https://stackoverflow.com/a/3804852/5536802
        public static bool HasEqualContent<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, IDictionary<TKey, TValue> otherDictionary)
        {
            return dictionary.Count == otherDictionary.Count && !dictionary.Except(otherDictionary).Any();
        }
    }
}
