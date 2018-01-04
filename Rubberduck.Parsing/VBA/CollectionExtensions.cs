using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA
{
    // See http://stackoverflow.com/questions/34362316/how-to-turn-icollectiont-into-ireadonlycollectiont
    public class ReadOnlyCollectionWrapper<T> : IReadOnlyCollection<T>
    {
        private readonly ICollection<T> _collection;
        public ReadOnlyCollectionWrapper(ICollection<T> collection)
        {
            _collection = collection;
        }

        public int Count => _collection.Count;

        public IEnumerator<T> GetEnumerator()
        {
            return _collection.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _collection.GetEnumerator();
        }
    }

    public static class CollectionExtensions
    {
        public static IReadOnlyCollection<T> AsReadOnly<T>(this ICollection<T> collection)
        {
            return new ReadOnlyCollectionWrapper<T>(collection);
        }
    }
}
