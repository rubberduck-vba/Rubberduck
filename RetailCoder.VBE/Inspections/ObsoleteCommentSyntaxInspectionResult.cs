using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.UI;
using Rubberduck.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ObsoleteCommentSyntaxInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ObsoleteCommentSyntaxInspectionResult(IInspection inspection, CommentNode comment) 
            : base(inspection, inspection.Description, comment)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new ReplaceCommentMarkerQuickFix(Context, QualifiedSelection, comment),
                new RemoveCommentQuickFix(Context, QualifiedSelection, comment), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class RemoveCommentQuickFix : CodeInspectionQuickFix
    {
        private readonly CommentNode _comment;

        public RemoveCommentQuickFix(ParserRuleContext context, QualifiedSelection selection, CommentNode comment)
            : base(context, selection, RubberduckUI.Inspections_RemoveComment)
        {
            _comment = comment;
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var content = module.get_Lines(Selection.Selection.StartLine, Selection.Selection.LineCount);

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

            if (_comment.QualifiedSelection.Selection.LineCount > 1)
            {
                module.DeleteLines(_comment.QualifiedSelection.Selection.StartLine, _comment.QualifiedSelection.Selection.LineCount);
            }

            module.ReplaceLine(_comment.QualifiedSelection.Selection.StartLine, code);
        }
    }

    public class ReplaceCommentMarkerQuickFix : CodeInspectionQuickFix
    {
        private readonly CommentNode _comment;

        public ReplaceCommentMarkerQuickFix(ParserRuleContext context, QualifiedSelection selection, CommentNode comment)
            : base(context, selection, RubberduckUI.Inspections_ReplaceRemWithSingleQuoteMarker)
        {
            _comment = comment;
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var content = module.get_Lines(Selection.Selection.StartLine, Selection.Selection.LineCount);

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

            var newContent = code + Tokens.CommentMarker + " " + _comment.CommentText;

            if (_comment.QualifiedSelection.Selection.LineCount > 1)
            {
                module.DeleteLines(_comment.QualifiedSelection.Selection.StartLine + 1, _comment.QualifiedSelection.Selection.LineCount);
            }

            module.ReplaceLine(Selection.Selection.StartLine, newContent);
        }
    }
}