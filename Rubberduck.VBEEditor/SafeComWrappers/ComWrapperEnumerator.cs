using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class ComWrapperEnumerator<TWrapperItem> : IEnumerator<TWrapperItem>
        where TWrapperItem : class
    {
        private readonly Func<object, TWrapperItem> _itemWrapper;
        private readonly IEnumerator _internal;

        public ComWrapperEnumerator(IEnumerable source, Func<object, TWrapperItem> itemWrapper)
        {
            _itemWrapper = itemWrapper;
            _internal = source?.GetEnumerator() ?? Enumerable.Empty<TWrapperItem>().GetEnumerator();
        }

        public void Dispose()
        {
            // nothing to dispose here
        }

        public bool MoveNext()
        {
            return _internal.MoveNext();
        }

        public void Reset()
        {
            _internal.Reset();
        }

        public TWrapperItem Current => _itemWrapper.Invoke(_internal.Current);

        object IEnumerator.Current => Current;
    }
}