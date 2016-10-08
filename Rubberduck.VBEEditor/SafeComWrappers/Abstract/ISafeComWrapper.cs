namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
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