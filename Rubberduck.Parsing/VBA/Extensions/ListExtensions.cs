using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA.Extensions
{
    public class ReadOnlyListWrapper<T> : IReadOnlyList<T>
    {
        private readonly IList<T> _list;
        public ReadOnlyListWrapper(IList<T> list)
        {
            _list = list;
        }

        public int Count => _list.Count;
        public IEnumerator<T> GetEnumerator() => _list.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => _list.GetEnumerator();
        public T this[int index] => _list[index];
    }

    public static class ListExtensions
    {
        public static IReadOnlyList<T> AsReadOnly<T>(this IList<T> list)
        {
            return new ReadOnlyListWrapper<T>(list);
        }
    }
}