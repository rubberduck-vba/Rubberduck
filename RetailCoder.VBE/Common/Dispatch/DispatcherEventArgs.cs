using System;
using Rubberduck.Parsing;

namespace Rubberduck.Common.Dispatch
{
    public class DispatcherEventArgs<T> : EventArgs, IDispatcherEventArgs<T>
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
