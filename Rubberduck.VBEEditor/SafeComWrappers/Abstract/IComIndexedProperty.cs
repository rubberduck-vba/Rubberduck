namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IComIndexedProperty<out TItem>
    {
        TItem this[object index] { get; }
    }
}