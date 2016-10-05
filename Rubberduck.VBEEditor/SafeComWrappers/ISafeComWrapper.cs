namespace Rubberduck.VBEditor.SafeComWrappers
{
    public interface ISafeComWrapper
    {
        /// <summary>
        /// Releases all COM objects.
        /// </summary>
        void Release();
    }
}