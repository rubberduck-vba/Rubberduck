using System;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class IntroduceLocalVariableForRedimQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public IntroduceLocalVariableForRedimQuickFix(RubberduckParserState state)
            : base(typeof(UndeclaredRedimVariableInspection))
        {
            _state = state;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        public override void Fix(IInspectionResult result)
        {
            VBAParser.RedimVariableDeclarationContext redimVariable = FindContext<VBAParser.RedimVariableDeclarationContext>(result.Target.Context);
            VBAParser.RedimStmtContext redimStatement = FindContext<VBAParser.RedimStmtContext>(redimVariable);
            var instruction = $"{Environment.NewLine}Dim {result.Target.IdentifierName}() As {result.Target.AsTypeName}{Environment.NewLine}";
            var rewriter = _state.GetRewriter(result.Target);
            rewriter.InsertBefore(redimStatement.Start.TokenIndex, instruction);
            // Remove the as type because we don't need that anymore.
            var asType = result.Target.AsTypeContext;
            if (asType != null)
            {
                rewriter.RemoveRange(asType.Start.TokenIndex - 1, asType.Stop.TokenIndex);
            }
            // Remove the type hint because we don't need that anymore.
            if (result.Target.HasTypeHint)
            {
                rewriter.RemoveRange(result.Target.Context.Stop.TokenIndex, result.Target.Context.Stop.TokenIndex);
            }
        }

        private T FindContext<T>(RuleContext context) where T : RuleContext
        {
            RuleContext temp = context;
            while (!(temp is T))
            {
                temp = temp.Parent;
            }
            return (T)temp;
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IntroduceLocalVariableQuickFix;
    }
}