namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ISafeComWrapper : INullObjectWrapper
    {
        //void Release(bool final = false);
    }

    public interface ISafeComWrapper<out T> : ISafeComWrapper
    {
        new T Target { get; }
    }

    public interface INullObjectWrapper
    {
        object Target { get; }
        bool IsWrappingNullReference { get; }
    }
}