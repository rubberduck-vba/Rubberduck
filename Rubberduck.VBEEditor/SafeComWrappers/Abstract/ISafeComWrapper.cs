namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ISafeComWrapper
    {
        object ComObject { get; }
        void Release();
        bool IsWrappingNullReference { get; }
    }
}