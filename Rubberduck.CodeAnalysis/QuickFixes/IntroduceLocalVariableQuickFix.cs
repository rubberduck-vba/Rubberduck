using System;
using System.Text.RegularExpressions;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class IntroduceLocalVariableQuickFix : QuickFixBase
    {
        public IntroduceLocalVariableQuickFix()
            : base(typeof(UndeclaredVariableInspection))
        {}

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var identifierContext = result.Target.Context;
            var enclosingStatmentContext = identifierContext.GetAncestor<VBAParser.BlockStmtContext>();
            var instruction = IdentifierDeclarationText(result.Target.IdentifierName, EndOfStatementText(enclosingStatmentContext), FrontPadding(enclosingStatmentContext));
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            rewriter.InsertBefore(enclosingStatmentContext.Start.TokenIndex, instruction);
        }

        private string EndOfStatementText(VBAParser.BlockStmtContext context)
        {
            if (!context.TryGetPrecedingContext<VBAParser.IndividualNonEOFEndOfStatementContext>(out var individualEndOfStmtContext))
            {
                return Environment.NewLine;
            }

            var endOfLine = individualEndOfStmtContext.endOfLine();

            if (endOfLine?.commentOrAnnotation() == null)
            {
                return individualEndOfStmtContext.GetText();
            }

            //There is a comment inside the preceding endOfLine, which we do not want to duplicate.
            var whitespaceContext = individualEndOfStmtContext.whiteSpace(0);
            return Environment.NewLine + (whitespaceContext?.GetText() ?? string.Empty);
        }

        private string FrontPadding(VBAParser.BlockStmtContext context)
        {
            var statementLabelContext = context.statementLabelDefinition();
            if (statementLabelContext == null)
            {
                return string.Empty;
            }

            var statementLabelTextAsWhitespace = ReplaceNonWhitespaceWithSpace(statementLabelContext.GetText());
            var whitespaceContext = context.whiteSpace();
            return statementLabelTextAsWhitespace + (whitespaceContext?.GetText() ?? string.Empty);
        }

        private string ReplaceNonWhitespaceWithSpace(string input)
        {
            if (input == null || input.Equals(string.Empty))
            {
                return string.Empty;
            }

            var pattern = @"[^\r\n\t ]";
            var replacement = " ";
            var regex = new Regex(pattern);
            return regex.Replace(input, replacement);
        }

        private string IdentifierDeclarationText(string identifierName, string endOfStatementText, string prefix)
        {
            return $"{prefix}Dim {identifierName} As Variant{endOfStatementText}";
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IntroduceLocalVariableQuickFix;
    }
}