using System;
using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class ComWrapperEnumerator<TComCollection, TWrapperItem> : IEnumerator<TWrapperItem>
        where TComCollection : IEnumerable
    {
        private readonly IEnumerator _internal;

        public ComWrapperEnumerator(TComCollection items)
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

        public TWrapperItem Current
        {
            get { return (TWrapperItem)Activator.CreateInstance(typeof(TWrapperItem), _internal.Current); }
        }

        object IEnumerator.Current
        {
            get { return Current; }
        }
    }
}