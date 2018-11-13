using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class ProjectApp : HostApplicationBase<Microsoft.Office.Interop.MSProject.Application>
    {
        public ProjectApp(IVBE vbe) : base(vbe, "MSProject", true) { }
    }
}
