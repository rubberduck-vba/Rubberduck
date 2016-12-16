using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ReplaceObsoleteCommentMarkerQuickFix : QuickFixBase
    {
        private readonly CommentNode _comment;

        public ReplaceObsoleteCommentMarkerQuickFix(ParserRuleContext context, QualifiedSelection selection, CommentNode comment)
            : base(context, selection, InspectionsUI.ReplaceCommentMarkerQuickFix)
        {
            _comment = comment;
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            {
                if (module.IsWrappingNullReference)
                {
                    return;
                }

                var content = module.GetLines(Selection.Selection.StartLine, Selection.Selection.LineCount);

                int markerPosition;
                if (!content.HasComment(out markerPosition))
                {
                    return;
                }

                var code = string.Empty;
                if (markerPosition > 0)
                {
                    code = content.Substring(0, markerPosition);
                }

                var newContent = code + Tokens.CommentMarker + " " + _comment.CommentText;

                if (_comment.QualifiedSelection.Selection.LineCount > 1)
                {
                    module.DeleteLines(_comment.QualifiedSelection.Selection.StartLine + 1, _comment.QualifiedSelection.Selection.LineCount);
                }

                module.ReplaceLine(Selection.Selection.StartLine, newContent);
            }
        }
    }
}