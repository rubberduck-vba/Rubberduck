using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class ExcelApp : HostApplicationBase<Microsoft.Office.Interop.Excel.Application>
    {
        public ExcelApp(IVBE vbe) : base(vbe, "Excel", true) { }

        public override IEnumerable<HostAutoMacro> AutoMacroIdentifiers => new HostAutoMacro[]
        {
            new HostAutoMacro(new[] {ComponentType.StandardModule}, true, null, "Auto_Open"),
            new HostAutoMacro(new[] {ComponentType.StandardModule}, true, null, "Auto_Close")
        };
    }
}
