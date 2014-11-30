using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitByRefParameterInspectionResult : CodeInspectionResultBase
    {
        public ImplicitByRefParameterInspectionResult(string inspection, Instruction instruction,
            CodeInspectionSeverity type, string message)
            : base(inspection, instruction, type, message)
        {
        }

        public override void QuickFix(VBE vbe)
        {
            throw new System.NotImplementedException();
        }
    }
}