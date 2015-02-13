using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspectionResult : CodeInspectionResultBase
    {
        public OptionBaseInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName)
            : base(inspection, type, new CommentNode("", new QualifiedSelection(qualifiedName, Selection.Empty)))
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return new Dictionary<string, Action<VBE>>(); // these fixes could break the code.
            /*
                new Dictionary<string, Action<VBE>>
                {
                    {"Remove Option statement", RemoveOptionStatement},
                    {"Specify Option Base 0", SpecifyOptionBaseZero}
                }; 
            */
        }

        private void SpecifyOptionBaseZero(VBE vbe)
        {
            RebaseAllArrayReferences(vbe);
        }

        private void RebaseAllArrayReferences(VBE vbe)
        {
            throw new NotImplementedException();
        }

        private void RemoveOptionStatement(VBE vbe)
        {
            var module = vbe.FindCodeModules(QualifiedName).SingleOrDefault();
            if (module == null)
            {
                return;
            }

            var selection = Comment.QualifiedSelection.Selection;
            
            // remove line continuations to compare against Context:
            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount)
                .Replace("\r\n", " ")
                .Replace("_", string.Empty);
            var originalInstruction = Comment.Comment;

            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = string.Empty;
            var newCodeLines = string.IsNullOrEmpty(newInstruction)
                ? string.Empty
                : originalCodeLines.Replace(originalInstruction, newInstruction);

            if (!string.IsNullOrEmpty(newCodeLines))
            {
                module.InsertLines(selection.StartLine, newCodeLines);
            }

            RebaseAllArrayReferences(vbe);
        }
    }
}