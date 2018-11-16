using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VisioApp : HostApplicationBase<Microsoft.Office.Interop.Visio.Application>
    {
        public VisioApp(IVBE vbe) : base(vbe, "Visio", true) { }
    }
}
