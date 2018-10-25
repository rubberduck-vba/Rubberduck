using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class PowerPointApp : HostApplicationBase<Microsoft.Office.Interop.PowerPoint.Application>
    {
        private readonly IVBE _vbe;

        public PowerPointApp(IVBE vbe) : base(vbe, "PowerPoint")
        {
            _vbe = vbe;
        }
    }
}
