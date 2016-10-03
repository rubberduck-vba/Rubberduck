using System;
using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class ComWrapperEnumerator<TCollection, TItem> : IEnumerator<TItem>
        where TCollection : IEnumerable
    {
        private readonly IEnumerator _internal;

        public ComWrapperEnumerator(TCollection items)
        {
            _internal = items.GetEnumerator();
        }

        public void Dispose()
        {
            var disposable = _internal as IDisposable;
            if (disposable == null)
            {
                return;
            }
            disposable.Dispose();
        }

        public bool MoveNext()
        {
            return _internal.MoveNext();
        }

        public void Reset()
        {
            _internal.Reset();
        }

        public TItem Current
        {
            get { return (TItem)Activator.CreateInstance(typeof(TItem), _internal.Current); }
        }

        object IEnumerator.Current
        {
            get { return Current; }
        }
    }
}