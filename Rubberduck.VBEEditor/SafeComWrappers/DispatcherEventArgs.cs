using System;

namespace Rubberduck.VBEditor.SafeComWrappers
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

    public class DispatcherRenamedEventArgs<T> : DispatcherEventArgs<T>
        where T : class
    {
        private readonly string _oldName;

        public DispatcherRenamedEventArgs(T item, string oldName)
            : base(item)
        {
            _oldName = oldName;
        }

        public string OldName { get { return _oldName; } }
    }
}
