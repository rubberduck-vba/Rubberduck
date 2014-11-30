using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

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
            int index;
            if (!Instruction.Line.Content.HasComment(out index)) return;
            
            var line = Instruction.Line.Content.Substring(0, index) + "'" + Instruction.Comment.Substring(ReservedKeywords.Rem.Length);
            location.CodeModule.ReplaceLine(location.Selection.StartLine, line);
        }
    }
}