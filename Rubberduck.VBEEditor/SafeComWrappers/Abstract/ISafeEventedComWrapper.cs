namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ISafeEventedComWrapper : ISafeComWrapper
    {
        void AttachEvents();
        void DetachEvents();
    }
}
