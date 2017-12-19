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
using Rubberduck.Parsing.Inspections.Resources;
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

        public override void Fix(IInspectionResult result)
        {
            var annotationText = $"'@Ignore {result.Inspection.AnnotationName}";

            var module = result.QualifiedSelection.QualifiedName.Component.CodeModule;
            var annotationLine = result.QualifiedSelection.Selection.StartLine;
            while (annotationLine != 1 && module.GetLines(annotationLine - 1, 1).EndsWith(" _"))
            {
                annotationLine--;
            }
            var codeLine = annotationLine == 1 ? string.Empty : module.GetLines(annotationLine - 1, 1);

            RuleContext treeRoot = result.Context;
            while (treeRoot.Parent != null)
            {
                treeRoot = treeRoot.Parent;
            }
            
            if (codeLine.HasComment(out var commentStart) && codeLine.Substring(commentStart).StartsWith("'@Ignore "))
            {
                var listener = new AnnotationListener();
                ParseTreeWalker.Default.Walk(listener, treeRoot);

                var annotationContext = listener.Contexts.Last(i => i.Start.TokenIndex <= result.Context.Start.TokenIndex);

                var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
                rewriter.InsertAfter(annotationContext.annotationName().Stop.TokenIndex, $" {result.Inspection.AnnotationName},");
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
                    var listener = new EOLListener();
                    ParseTreeWalker.Default.Walk(listener, treeRoot);

                    // we subtract 2 here to get the insertion index to A) account for VBE's one-based indexing
                    // and B) to get the newline token that introduces that line
                    var eolContext = listener.Contexts.OrderBy(o => o.Start.TokenIndex).ElementAt(annotationLine - 2);
                    insertIndex = eolContext.Start.TokenIndex;

                    annotationText = Environment.NewLine + annotationText;
                }

                var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
                rewriter.InsertBefore(insertIndex, annotationText);
            }
        }

        public override string Description(IInspectionResult result) => InspectionsUI.IgnoreOnce;

        private class AnnotationListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.AnnotationContext> _contexts = new List<VBAParser.AnnotationContext>();
            public IEnumerable<VBAParser.AnnotationContext> Contexts => _contexts;

            public override void ExitAnnotation([NotNull] VBAParser.AnnotationContext context)
            {
                if (context.annotationName().GetText() == Annotations.IgnoreInspection)
                {
                    _contexts.Add(context);
                }
            }
        }

        private class EOLListener : VBAParserBaseListener
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
