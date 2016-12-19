using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveCommentQuickFix : QuickFixBase
    {
        private readonly CommentNode _comment;

        public RemoveCommentQuickFix(ParserRuleContext context, QualifiedSelection selection, CommentNode comment)
            : base(context, selection, InspectionsUI.RemoveCommentQuickFix)
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
                    code = content.Substring(0, markerPosition).TrimEnd();
                }

                if (_comment.QualifiedSelection.Selection.LineCount > 1)
                {
                    module.DeleteLines(_comment.QualifiedSelection.Selection.StartLine, _comment.QualifiedSelection.Selection.LineCount);
                }

                module.ReplaceLine(_comment.QualifiedSelection.Selection.StartLine, code);
            }
        }
    }
}