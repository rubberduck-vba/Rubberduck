namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IEventSource<out TEventSource>
    {
        TEventSource EventSource { get; }
    }
}
