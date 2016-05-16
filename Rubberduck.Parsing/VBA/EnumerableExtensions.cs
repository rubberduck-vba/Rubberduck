using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.VBA
{
    public static class EnumerableExtensions
    {
        public static IEnumerable<T> DistinctBy<T, TKey>(this IEnumerable<T> source, Func<T, TKey> keySelector)
        {
            if (source == null)
            {
                throw new ArgumentNullException("source");
            }
            if (keySelector == null)
            {
                throw new ArgumentNullException("keySelector");
            }

            var hashSet = new HashSet<TKey>();
            return source.Where(item => hashSet.Add(keySelector(item)));
        }
    }
}