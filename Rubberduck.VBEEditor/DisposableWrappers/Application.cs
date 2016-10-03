namespace Rubberduck.VBEditor.DisposableWrappers
{
    public class Application : SafeComWrapper<Microsoft.Vbe.Interop.Application>
    {
        public Application(Microsoft.Vbe.Interop.Application application)
            :base(application)
        {
        }

        public string Version { get { return InvokeResult(() => ComObject.Version); } }
    }
}