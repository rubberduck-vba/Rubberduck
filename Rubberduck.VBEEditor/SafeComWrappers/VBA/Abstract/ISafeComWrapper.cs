namespace Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract
{
    /// <summary>
    /// 
    /// </summary>
    public interface ISafeComWrapper
    {
        object ComObject { get; }
        bool IsWrappingNullReference { get; }
        void Release();
    }
}