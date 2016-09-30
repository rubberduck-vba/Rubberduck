namespace Rubberduck.VBEditor.DisposableWrappers
{
    internal interface IWrapper<out T>
    {
        T WrappedInteropObject { get; }
    }
}