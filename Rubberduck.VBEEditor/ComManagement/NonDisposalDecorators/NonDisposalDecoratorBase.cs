using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement.NonDisposalDecorators
{
    /// <summary>
    /// Decorator for SafeComWrappers to safely hand out references that must not be disposed
    /// </summary>
    /// <typeparam name="T">Concrete type of the safe com wrapper to decorate</typeparam>
    public class NonDisposalDecoratorBase<T> : ISafeComWrapper
        where T :ISafeComWrapper
    {
        protected readonly T WrappedItem;

        public NonDisposalDecoratorBase(T wrappedItem)
        {
            WrappedItem = wrappedItem;
        }

        public object Target => WrappedItem.Target;
        public bool IsWrappingNullReference => WrappedItem.IsWrappingNullReference;
        public void Dispose()
        {
            //Do nothing
        }
    }
}