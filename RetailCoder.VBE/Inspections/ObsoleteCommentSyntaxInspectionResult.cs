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
        public ObsoleteCommentSyntaxInspectionResult(string inspection, SyntaxTreeNode node, CodeInspectionSeverity type) 
            : base(inspection, node, type)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return !Handled
                ? new Dictionary<string, Action<VBE>>
                    {
                        {"Replace Rem reserved keyword with single quote", ReplaceWithSingleQuote},
                        {"Remove comment", RemoveComment}
                    }
                : new Dictionary<string, Action<VBE>>();
        }

        private void ReplaceWithSingleQuote(VBE vbe)
        {
            var instruction = Node.Instruction;
            var location = vbe.FindInstruction(instruction);
            int index;
            if (!instruction.Line.Content.HasComment(out index)) return;
            
            var line = instruction.Line.Content.Substring(0, index) + "'" + instruction.Comment.Substring(ReservedKeywords.Rem.Length);
            location.CodeModule.ReplaceLine(location.Selection.StartLine, line);

            Handled = true;
        }

        private void RemoveComment(VBE vbe)
        {
            var instruction = Node.Instruction;
            var location = vbe.FindInstruction(instruction);
            int index;
            if (!instruction.Line.Content.HasComment(out index)) return;

            var line = instruction.Line.Content.Substring(0, index);
            location.CodeModule.ReplaceLine(location.Selection.StartLine, line);

            Handled = true;
        }
    }
}