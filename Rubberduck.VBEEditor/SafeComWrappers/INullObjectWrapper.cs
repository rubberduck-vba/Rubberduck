namespace Rubberduck.VBEditor.SafeComWrappers
{
    public interface INullObjectWrapper
    {
        object Target { get; }
        bool IsWrappingNullReference { get; }
    }
}
