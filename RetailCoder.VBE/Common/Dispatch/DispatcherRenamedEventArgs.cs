namespace Rubberduck.Common.Dispatch
{
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