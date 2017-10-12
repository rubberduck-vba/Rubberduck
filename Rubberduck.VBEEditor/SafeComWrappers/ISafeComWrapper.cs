namespace Rubberduck.VBEditor.SafeComWrappers
{
    public interface ISafeComWrapper : INullObjectWrapper
    {
        //void Release(bool final = false);
    }

    public interface ISafeComWrapper<out T> : ISafeComWrapper
    {
        new T Target { get; }
    }
}