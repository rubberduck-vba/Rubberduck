using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class PowerPointApp : HostApplicationBase<Microsoft.Office.Interop.PowerPoint.Application>
    {
        public PowerPointApp(IVBE vbe) : base(vbe, "PowerPoint") { }

        public override IEnumerable<HostAutoMacro> AutoMacroIdentifiers => new HostAutoMacro[]
        {
            // Technically, those are only run only if it's an add-in, not opened directly
            new HostAutoMacro(new[] {ComponentType.StandardModule}, false, null, "Auto_Open"),
            new HostAutoMacro(new[] {ComponentType.StandardModule}, false, null, "Auto_Close")
        };
    }
}
