using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ObsoleteCommentSyntaxInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteCommentSyntaxInspectionResult(string inspection, Instruction instruction, CodeInspectionSeverity type) 
            : base(inspection, instruction, type)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return !Handled
                ? new Dictionary<string, Action<VBE>>
                {
                    {"Replace Rem reserved keyword with single quote.", ReplaceWithSingleQuote},
                    {"Remove comment.", RemoveComment}
                }
                : new Dictionary<string, Action<VBE>>();
        }

        private void ReplaceWithSingleQuote(VBE vbe)
        {
            var location = vbe.FindInstruction(Instruction);
            int index;
            if (!Instruction.Line.Content.HasComment(out index)) return;
            
            var line = Instruction.Line.Content.Substring(0, index) + "'" + Instruction.Comment.Substring(ReservedKeywords.Rem.Length);
            location.CodeModule.ReplaceLine(location.Selection.StartLine, line);
        }

        private void RemoveComment(VBE vbe)
        {
            var location = vbe.FindInstruction(Instruction);
            int index;
            if (!Instruction.Line.Content.HasComment(out index)) return;

            var line = Instruction.Line.Content.Substring(0, index);
            location.CodeModule.ReplaceLine(location.Selection.StartLine, line);

            Handled = true;
        }
    }
}