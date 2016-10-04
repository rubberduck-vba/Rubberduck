using System;
using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class ComWrapperEnumerator<TWrapperItem> : IEnumerator<TWrapperItem>
        where TWrapperItem : class
    {
        private readonly IEnumerator _internal;

        public ComWrapperEnumerator(IEnumerable source)
        {
            _internal = source.GetEnumerator();
        }

        public void Dispose()
        {
            var disposable = _internal as IDisposable;
            if (disposable != null)
            {
                // COM enumerator won't dispose
                disposable.Dispose();
            }
            else
            {
                // COM enumerator won't cast to __ComObject either
                // Marshal.ReleaseComObject(_internal);
            }
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