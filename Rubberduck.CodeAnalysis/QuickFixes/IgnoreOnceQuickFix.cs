using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class IgnoreOnceQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public IgnoreOnceQuickFix(RubberduckParserState state, IEnumerable<IInspection> inspections)
            : base(inspections.Select(s => s.GetType()).Where(i => i.CustomAttributes.All(a => a.AttributeType != typeof(CannotAnnotateAttribute))).ToArray())
        {
            _state = state;
        }

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var annotationText = $"'@Ignore {result.Inspection.AnnotationName}";

            int annotationLine;
            //TODO: Make this use the parse tree instead of the code module.
            var component = _state.ProjectsProvider.Component(result.QualifiedSelection.QualifiedName);
            using (var module = component.CodeModule)
            {
                annotationLine = result.QualifiedSelection.Selection.StartLine;
                while (annotationLine != 1 && module.GetLines(annotationLine - 1, 1).EndsWith(" _"))
                {
                    annotationLine--;
                }
            }

            RuleContext treeRoot = result.Context;
            while (treeRoot.Parent != null)
            {
                treeRoot = treeRoot.Parent;
            }

            var listener = new CommentOrAnnotationListener();
            ParseTreeWalker.Default.Walk(listener, treeRoot);
            var commentContext = listener.Contexts.LastOrDefault(i => i.Stop.TokenIndex <= result.Context.Start.TokenIndex);
            var commented = commentContext?.Stop.Line + 1 == annotationLine;

            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            if (commented)
            {
                var annotation = commentContext.annotationList()?.annotation(0);
                if (annotation != null && annotation.GetText().StartsWith("Ignore"))
                {
                    rewriter.InsertAfter(annotation.annotationName().Stop.TokenIndex, $" {result.Inspection.AnnotationName},");
                }
                else
                {
                    var indent = new string(Enumerable.Repeat(' ', commentContext.Start.Column).ToArray());
                    rewriter.InsertAfter(commentContext.Stop.TokenIndex, $"{indent}{annotationText}{Environment.NewLine}");
                }
            }
            else
            {
                int insertIndex;

                // this value is used when the annotation should be on line 1--we need to insert before token index 0
                if (annotationLine == 1)
                {
                    insertIndex = 0;
                    annotationText += Environment.NewLine;
                }
                else
                {
                    var eol = new EndOfLineListener();
                    ParseTreeWalker.Default.Walk(eol, treeRoot);

                    // we subtract 2 here to get the insertion index to A) account for VBE's one-based indexing
                    // and B) to get the newline token that introduces that line
                    var eolContext = eol.Contexts.OrderBy(o => o.Start.TokenIndex).ElementAt(annotationLine - 2);
                    insertIndex = eolContext.Start.TokenIndex;

                    annotationText = Environment.NewLine + annotationText;
                }

                rewriter.InsertBefore(insertIndex, annotationText);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IgnoreOnce;

        private class CommentOrAnnotationListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.CommentOrAnnotationContext> _contexts = new List<VBAParser.CommentOrAnnotationContext>();
            public IEnumerable<VBAParser.CommentOrAnnotationContext> Contexts => _contexts;

            public override void ExitCommentOrAnnotation([NotNull] VBAParser.CommentOrAnnotationContext context)
            {
                _contexts.Add(context);
            }
        }

        private class EndOfLineListener : VBAParserBaseListener
        {
            private readonly IList<ParserRuleContext> _contexts = new List<ParserRuleContext>();
            public IEnumerable<ParserRuleContext> Contexts => _contexts;

            public override void ExitWhiteSpace([NotNull] VBAParser.WhiteSpaceContext context)
            {
                if (context.GetText().Contains(Environment.NewLine))
                {
                    _contexts.Add(context);
                }
            }

            public override void ExitEndOfLine([NotNull] VBAParser.EndOfLineContext context)
            {
                _contexts.Add(context);
            }
        }
    }
}
