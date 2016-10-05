namespace Rubberduck.VBEditor.DisposableWrappers
{
    public interface ISafeComWrapper
    {
        /// <summary>
        /// Releases all COM objects.
        /// </summary>
        void Release();
    }
}