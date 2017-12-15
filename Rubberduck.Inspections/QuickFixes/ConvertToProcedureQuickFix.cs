using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ConvertToProcedureQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public ConvertToProcedureQuickFix(RubberduckParserState state)
            : base(typeof(NonReturningFunctionInspection), typeof(FunctionReturnValueNotUsedInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            switch (result.Context)
            {
                case VBAParser.FunctionStmtContext functionContext:
                    ConvertFunction(result, functionContext);
                    break;
                case VBAParser.PropertyGetStmtContext propertyGetContext:
                    ConvertPropertyGet(result, propertyGetContext);
                    break;
            }
        }

        private void ConvertFunction(IInspectionResult result, VBAParser.FunctionStmtContext functionContext)
        {
            var rewriter = _state.GetRewriter(result.Target);

            var asTypeContext = ParserRuleContextHelper.GetChild<VBAParser.AsTypeClauseContext>(functionContext);
            if (asTypeContext != null)
            {
                rewriter.Remove(asTypeContext);
                rewriter.Remove(functionContext.children.ElementAt(functionContext.children.IndexOf(asTypeContext) - 1) as ParserRuleContext);
            }

            if (result.Target.TypeHint != null)
            {
                rewriter.Remove(ParserRuleContextHelper.GetDescendent<VBAParser.TypeHintContext>(functionContext));
            }

            rewriter.Replace(functionContext.FUNCTION(), Tokens.Sub);
            rewriter.Replace(functionContext.END_FUNCTION(), "End Sub");

            foreach (var returnStatement in GetReturnStatements(result.Target))
            {
                rewriter.Remove(returnStatement);
            }
        }

        private void ConvertPropertyGet(IInspectionResult result, VBAParser.PropertyGetStmtContext propertyGetContext)
        {
            var rewriter = _state.GetRewriter(result.Target);

            var asTypeContext = ParserRuleContextHelper.GetChild<VBAParser.AsTypeClauseContext>(propertyGetContext);
            if (asTypeContext != null)
            {
                rewriter.Remove(asTypeContext);
                rewriter.Remove(propertyGetContext.children.ElementAt(propertyGetContext.children.IndexOf(asTypeContext) - 1) as ParserRuleContext);
            }

            if (result.Target.TypeHint != null)
            {
                rewriter.Remove(ParserRuleContextHelper.GetDescendent<VBAParser.TypeHintContext>(propertyGetContext));
            }

            rewriter.Replace(propertyGetContext.PROPERTY_GET(), Tokens.Sub);
            rewriter.Replace(propertyGetContext.END_PROPERTY(), "End Sub");

            foreach (var returnStatement in GetReturnStatements(result.Target))
            {
                rewriter.Remove(returnStatement);
            }
        }

        public override string Description(IInspectionResult result) =>InspectionsUI.ConvertFunctionToProcedureQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => false;

        private IEnumerable<ParserRuleContext> GetReturnStatements(Declaration declaration)
        {
            return declaration.References
                .Where(usage => IsReturnStatement(declaration, usage))
                .Select(usage => usage.Context.Parent)
                .Cast<ParserRuleContext>();
        }

        private bool IsReturnStatement(Declaration declaration, IdentifierReference assignment)
        {
            return assignment.ParentScoping.Equals(declaration) && assignment.Declaration.Equals(declaration);
        }
    }
}
