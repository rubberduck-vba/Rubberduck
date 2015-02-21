using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.Inspections
{
    public class ObsoleteCommentSyntaxInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteCommentSyntaxInspectionResult(string inspection, CodeInspectionSeverity type, CommentNode comment) 
            : base(inspection, type, comment)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return
                new Dictionary<string, Action<VBE>>
                {
                    {"Replace 'Rem' usage with a single-quote comment marker", ReplaceWithSingleQuote},
                    {"Remove comment", RemoveComment}
                };
        }

        private void ReplaceWithSingleQuote(VBE vbe)
        {
            var module = vbe.FindCodeModules(QualifiedName).FirstOrDefault();
            if (module == null)
            {
                return;
            }

            var content = module.get_Lines(QualifiedSelection.Selection.StartLine, QualifiedSelection.Selection.LineCount);

            int markerPosition;
            if (!content.HasComment(out markerPosition))
            {
                return;
            }

            var code = string.Empty;
            if (markerPosition > 0)
            {
                code = content.Substring(0, markerPosition - 1);
            }

            var newContent = code + Tokens.CommentMarker + " " + Comment.CommentText;

            if (Comment.QualifiedSelection.Selection.LineCount > 1)
            {
                module.DeleteLines(Comment.QualifiedSelection.Selection.StartLine + 1, Comment.QualifiedSelection.Selection.LineCount);
            }

            module.ReplaceLine(QualifiedSelection.Selection.StartLine, newContent);
        }

        private void RemoveComment(VBE vbe)
        {
            var module = vbe.FindCodeModules(QualifiedName).FirstOrDefault();
            if (module == null)
            {
                return;
            }

            var content = module.get_Lines(QualifiedSelection.Selection.StartLine, QualifiedSelection.Selection.LineCount);

            int markerPosition;
            if (!content.HasComment(out markerPosition))
            {
                return;
            }

            var code = string.Empty;
            if (markerPosition > 0)
            {
                code = content.Substring(0, markerPosition - 1);
            }

            if (Comment.QualifiedSelection.Selection.LineCount > 1)
            {
                module.DeleteLines(Comment.QualifiedSelection.Selection.StartLine, Comment.QualifiedSelection.Selection.LineCount);
            }

            if (!string.IsNullOrEmpty(code))
            {
                module.ReplaceLine(Comment.QualifiedSelection.Selection.StartLine, code);
            }
        }
    }
}