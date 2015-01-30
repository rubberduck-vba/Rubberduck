using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ObsoleteCommentSyntaxInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteCommentSyntaxInspectionResult(string inspection, ParserRuleContext context, CodeInspectionSeverity type, string project, string module) 
            : base(inspection, context, type, project, module)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return
                new Dictionary<string, Action<VBE>>
                {
                    {"Replace Rem reserved keyword with single quote", ReplaceWithSingleQuote},
                    {"Remove comment", RemoveComment}
                };
        }

        private void ReplaceWithSingleQuote(VBE vbe)
        {
            //var instruction = Context.Instruction;
            //var location = vbe.FindInstruction(instruction);
            //int index;
            //if (!instruction.Line.Content.HasComment(out index)) return;
            
            //var line = instruction.Line.Content.Substring(0, index) + "'" + instruction.Comment.Substring(ReservedKeywords.Rem.Length);
            //location.CodeModule.ReplaceLine(location.Selection.StartLine, line);
        }

        private void RemoveComment(VBE vbe)
        {
            //var instruction = Context.Instruction;
            //var location = vbe.FindInstruction(instruction);
            //int index;
            //if (!instruction.Line.Content.HasComment(out index)) return;

            //var line = instruction.Line.Content.Substring(0, index);
            //location.CodeModule.ReplaceLine(location.Selection.StartLine, line);
        }
    }
}