using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA.Extensions
{

    // See http://stackoverflow.com/questions/34362316/how-to-turn-icollectiont-into-ireadonlycollectiont
    public class ReadOnlyCollectionWrapper<T> : IReadOnlyCollection<T>
    {
        private ICollection<T> collection;
        public ReadOnlyCollectionWrapper(ICollection<T> collection)
        {
            this.collection = collection;
        }

        public int Count
        {
            get { return collection.Count; }
        }

        public IEnumerator<T> GetEnumerator()
        {
            return collection.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return collection.GetEnumerator();
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
