using System;
using System.Text.RegularExpressions;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Introduces a local Variant variable for an otherwise undeclared identifier.
    /// </summary>
    /// <inspections>
    /// <inspection name="UndeclaredVariableInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     value = 42
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Variant
    ///     value = 42
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class IntroduceLocalVariableQuickFix : QuickFixBase
    {
        public IntroduceLocalVariableQuickFix()
            : base(typeof(UndeclaredVariableInspection))
        {}

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var identifierContext = result.Target.Context;
            var enclosingStatementContext = identifierContext.GetAncestor<VBAParser.BlockStmtContext>();
            var instruction = IdentifierDeclarationText(result.Target.IdentifierName, EndOfStatementText(enclosingStatementContext), FrontPadding(enclosingStatementContext));
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            rewriter.InsertBefore(enclosingStatementContext.Start.TokenIndex, instruction);
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