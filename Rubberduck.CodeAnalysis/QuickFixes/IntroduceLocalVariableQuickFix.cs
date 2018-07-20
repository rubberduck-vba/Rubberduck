using System;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class IntroduceLocalVariableQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public IntroduceLocalVariableQuickFix(RubberduckParserState state)
            : base(typeof(UndeclaredVariableInspection))
        {
            _state = state;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        public override void Fix(IInspectionResult result)
        {
            var identifierContext = result.Target.Context;
            var enclosingStatmentContext = identifierContext.GetAncestor<VBAParser.BlockStmtContext>();
            var instruction = IdentifierDeclarationText(result.Target.IdentifierName, EndOfStatementText(enclosingStatmentContext));
            _state.GetRewriter(result.Target).InsertBefore(enclosingStatmentContext.Start.TokenIndex, instruction);
        }

        private string EndOfStatementText(VBAParser.BlockStmtContext context)
        {
            if (context.TryGetPrecedingContext<VBAParser.EndOfStatementContext>(out var endOfStmtContext))
            {
                return endOfStmtContext.GetText();
            }

            return Environment.NewLine;
        }

        private string IdentifierDeclarationText(string identifierName, string endOfStatementText)
        {
            return $"Dim {identifierName} As Variant{endOfStatementText}";
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IntroduceLocalVariableQuickFix;
    }
}