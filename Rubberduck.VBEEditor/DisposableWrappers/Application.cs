namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class Application : WrapperBase<Microsoft.Vbe.Interop.Application>
    {
        public Application(Microsoft.Vbe.Interop.Application application)
            :base(application)
        {
        }

        public string Version
        {
            get
            {
                ThrowIfDisposed();
                return InvokeMemberValue(() => Item.Version);
            }
        }
    }
}