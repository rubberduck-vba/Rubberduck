using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    public class ObsoleteCommentSyntaxInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteCommentSyntaxInspectionResult(string inspection, Instruction instruction, CodeInspectionSeverity type, string message) 
            : base(inspection, instruction, type, message)
        {
        }

        public override void QuickFix(VBE vbe)
        {
            var location = vbe.FindInstruction(Instruction);
            location.CodeModule.ReplaceLine(location.Selection.StartLine, "' " + Instruction.Comment);
        }
    }
}