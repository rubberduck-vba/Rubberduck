using System;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class DispatcherEventArgs<T> : EventArgs
        where T : class
    {
        public DispatcherEventArgs(T item)
        {
            Item = item;
        }

        public T Item { get; }
    }

    public class DispatcherRenamedEventArgs<T> : DispatcherEventArgs<T>
        where T : class
    {
        public DispatcherRenamedEventArgs(T item, string oldName)
            : base(item)
        {
            OldName = oldName;
        }

        public string OldName { get; }
    }
}
