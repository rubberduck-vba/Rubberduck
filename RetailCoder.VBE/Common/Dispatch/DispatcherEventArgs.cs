using System;

namespace Rubberduck.Common.Dispatch
{
    public class DispatcherEventArgs<T> : EventArgs 
        where T : class
    {
        private readonly T _item;

        public DispatcherEventArgs(T item)
        {
            _item = item;
        }

        public T Item { get { return _item; } }
    }
}
