namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IReleasableCOMObject
    {
        void Release(bool final = false);
    }
}
