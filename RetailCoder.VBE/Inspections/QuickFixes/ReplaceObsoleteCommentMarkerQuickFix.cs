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
        public ReplaceObsoleteCommentMarkerQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveObsoleteStatementQuickFix)
        { }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;

            if (module.IsWrappingNullReference)
            {
                return;
            }
            var comment = Context.GetText();
            var start = Context.Start.Line;           
            var commentLine = module.GetLines(start, 1);
            var newComment = commentLine.Substring(0, Context.Start.Column) +
                             Tokens.CommentMarker +
                             comment.Substring(Tokens.Rem.Length, comment.Length - Tokens.Rem.Length);

            var lines = Selection.Selection.LineCount;
            if (lines > 1)
            {
                module.DeleteLines(start + 1, lines - 1);
            }
            module.ReplaceLine(start, newComment);
        }
    }
}