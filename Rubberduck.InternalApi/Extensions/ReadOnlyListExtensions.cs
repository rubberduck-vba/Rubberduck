using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.InternalApi.Extensions
{
    public static class ReadOnlyListExtensions
    {
        public static int FindIndex<T>(this IReadOnlyList<T> source, Predicate<T> predicate)
        {
            var (elem, index) = source.Select<T, (T item, int i)>((item, i) => (item, i)).FirstOrDefault(tpl => predicate(tpl.item));

            if (index > 0 || index == 0 && predicate(source[0]))
            {
                return index;
            }

            return -1;
        }

        public static int IndexOf<T>(this IReadOnlyList<T> source, T elem)
        {
            return source.FindIndex(item => elem.Equals(item));
        }
    }
}