using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ConvertToProcedureQuickFix : QuickFixBase
    {
        public ConvertToProcedureQuickFix()
            : base(typeof(NonReturningFunctionInspection), typeof(FunctionReturnValueNotUsedInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);

            switch (result.Context)
            {
                case VBAParser.FunctionStmtContext functionContext:
                    ConvertFunction(result, functionContext, rewriter);
                    break;
                case VBAParser.PropertyGetStmtContext propertyGetContext:
                    ConvertPropertyGet(result, propertyGetContext, rewriter);
                    break;
            }
        }

        private void ConvertFunction(IInspectionResult result, VBAParser.FunctionStmtContext functionContext, IModuleRewriter rewriter)
        {
            var asTypeContext = functionContext.GetChild<VBAParser.AsTypeClauseContext>();
            if (asTypeContext != null)
            {
                rewriter.Remove(asTypeContext);
                rewriter.Remove(functionContext.children.ElementAt(functionContext.children.IndexOf(asTypeContext) - 1) as ParserRuleContext);
            }

            if (result.Target.TypeHint != null)
            {
                rewriter.Remove(functionContext.GetDescendent<VBAParser.TypeHintContext>());
            }

            rewriter.Replace(functionContext.FUNCTION(), Tokens.Sub);
            rewriter.Replace(functionContext.END_FUNCTION(), "End Sub");

            foreach (var returnStatement in GetReturnStatements(result.Target))
            {
                rewriter.Remove(returnStatement);
            }
        }

        private void ConvertPropertyGet(IInspectionResult result, VBAParser.PropertyGetStmtContext propertyGetContext, IModuleRewriter rewriter)
        {
            var asTypeContext = propertyGetContext.GetChild<VBAParser.AsTypeClauseContext>();
            if (asTypeContext != null)
            {
                rewriter.Remove(asTypeContext);
                rewriter.Remove(propertyGetContext.children.ElementAt(propertyGetContext.children.IndexOf(asTypeContext) - 1) as ParserRuleContext);
            }

            if (result.Target.TypeHint != null)
            {
                rewriter.Remove(propertyGetContext.GetDescendent<VBAParser.TypeHintContext>());
            }

            rewriter.Replace(propertyGetContext.PROPERTY_GET(), Tokens.Sub);
            rewriter.Replace(propertyGetContext.END_PROPERTY(), "End Sub");

            foreach (var returnStatement in GetReturnStatements(result.Target))
            {
                rewriter.Remove(returnStatement);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ConvertFunctionToProcedureQuickFix;

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
