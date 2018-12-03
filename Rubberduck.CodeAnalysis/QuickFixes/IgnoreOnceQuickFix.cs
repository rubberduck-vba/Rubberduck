using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

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
            if (result.Target?.DeclarationType.HasFlag(DeclarationType.Module) ?? false)
            {
                FixModule(result, rewriteSession);
            }
            else
            {
                FixNonModule(result, rewriteSession);
            }
        }

        private void FixNonModule(IInspectionResult result, IRewriteSession rewriteSession)
        {
            int insertionIndex;
            string insertText;
            var annotationText = $"'@Ignore {result.Inspection.AnnotationName}";

            var module = result.QualifiedSelection.QualifiedName;
            var parseTree = _state.GetParseTree(module, CodeKind.CodePaneCode);
            var eolListener = new EndOfLineListener();
            ParseTreeWalker.Default.Walk(eolListener, parseTree);
            var previousEol = eolListener.Contexts
                .OrderBy(eol => eol.Start.TokenIndex)
                .LastOrDefault(eol => eol.Start.Line < result.QualifiedSelection.Selection.StartLine);

            var rewriter = rewriteSession.CheckOutModuleRewriter(module);

            if (previousEol == null)
            {
                // The context to get annotated is on the first line; we need to insert before token index 0.
                insertionIndex = 0;
                insertText = annotationText + Environment.NewLine;
                rewriter.InsertBefore(insertionIndex, insertText);
                return;
            }

            var commentContext = previousEol.commentOrAnnotation();
            if (commentContext == null)
            {
                insertionIndex = previousEol.Start.TokenIndex;
                var indent = WhitespaceAfter(previousEol);
                insertText = $"{Environment.NewLine}{indent}{annotationText}";
                rewriter.InsertBefore(insertionIndex, insertText);
                return;
            }

            var ignoreAnnotation = commentContext.annotationList()?.annotation()
                .FirstOrDefault(annotationContext => annotationContext.annotationName().GetText() == AnnotationType.Ignore.ToString());
            if (ignoreAnnotation == null)
            {
                insertionIndex = commentContext.Stop.TokenIndex;
                var indent = WhitespaceAfter(previousEol);
                insertText = $"{indent}{annotationText}{Environment.NewLine}";
                rewriter.InsertAfter(insertionIndex, insertText);
                return;
            }

            insertionIndex = ignoreAnnotation.annotationName().Stop.TokenIndex;
            insertText = $" {result.Inspection.AnnotationName},";
            rewriter.InsertAfter(insertionIndex, insertText);
        }

        private static string WhitespaceAfter(VBAParser.EndOfLineContext endOfLine)
        {
            var individualEndOfStatement = (VBAParser.IndividualNonEOFEndOfStatementContext) endOfLine.Parent;
            var whiteSpaceOnNextLine = individualEndOfStatement.whiteSpace(0);
            return whiteSpaceOnNextLine != null
                ? whiteSpaceOnNextLine.GetText()
                : string.Empty;
        }

        private void FixModule(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var module = result.QualifiedSelection.QualifiedName;
            var moduleAnnotations = _state.GetModuleAnnotations(module);
            var firstIgnoreModuleAnnotation = moduleAnnotations
                .Where(annotation => annotation.AnnotationType == AnnotationType.IgnoreModule)
                .OrderBy(annotation => annotation.Context.Start.TokenIndex)
                .FirstOrDefault();

            var rewriter = rewriteSession.CheckOutModuleRewriter(module);

            int insertionIndex;
            string insertText;

            if (firstIgnoreModuleAnnotation == null)
            {
                insertionIndex = 0;
                insertText = $"'@IgnoreModule {result.Inspection.AnnotationName}{Environment.NewLine}";
                rewriter.InsertBefore(insertionIndex, insertText);
                return;
            }

            insertionIndex = firstIgnoreModuleAnnotation.Context.annotationName().Stop.TokenIndex;
            insertText = $" {result.Inspection.AnnotationName},";
            rewriter.InsertAfter(insertionIndex, insertText);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IgnoreOnce;

        private class EndOfLineListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.EndOfLineContext> _contexts = new List<VBAParser.EndOfLineContext>();
            public IEnumerable<VBAParser.EndOfLineContext> Contexts => _contexts;

            public override void ExitEndOfLine([NotNull] VBAParser.EndOfLineContext context)
            {
                _contexts.Add(context);
            }
        }
    }
}
