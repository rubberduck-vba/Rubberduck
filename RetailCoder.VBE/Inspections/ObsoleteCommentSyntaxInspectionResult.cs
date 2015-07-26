using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.UI;
using Rubberduck.VBA;

namespace Rubberduck.Inspections
{
    public class ObsoleteCommentSyntaxInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteCommentSyntaxInspectionResult(string inspection, CodeInspectionSeverity type, CommentNode comment) 
            : base(inspection, type, comment)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return
                new Dictionary<string, Action>
                {
                    {RubberduckUI.Inspections_ReplaceRemWithSingleQuoteMarker, ReplaceWithSingleQuote},
                    {RubberduckUI.Inspections_RemoveComment, RemoveComment}
                };
        }

        private void ReplaceWithSingleQuote()
        {
            var module = QualifiedName.Component.CodeModule;
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

        private void RemoveComment()
        {
            var module = QualifiedName.Component.CodeModule;
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

            module.ReplaceLine(Comment.QualifiedSelection.Selection.StartLine, code);
        }
    }
}