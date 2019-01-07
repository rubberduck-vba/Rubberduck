using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveTypeHintsQuickFix : QuickFixBase
    {
        public RemoveTypeHintsQuickFix()
            : base(typeof(ObsoleteTypeHintInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!string.IsNullOrWhiteSpace(result.Target.TypeHint))
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
                var typeHintContext = result.Context.GetDescendent<VBAParser.TypeHintContext>();

                rewriter.Remove(typeHintContext);

                var asTypeClause = ' ' + Tokens.As + ' ' + SymbolList.TypeHintToTypeName[result.Target.TypeHint];
                switch (result.Target.DeclarationType)
                {
                    case DeclarationType.Variable:
                        var variableContext = (VBAParser.VariableSubStmtContext) result.Target.Context;
                        rewriter.InsertAfter(variableContext.identifier().Stop.TokenIndex, asTypeClause);
                        break;
                    case DeclarationType.Constant:
                        var constantContext = (VBAParser.ConstSubStmtContext) result.Target.Context;
                        rewriter.InsertAfter(constantContext.identifier().Stop.TokenIndex, asTypeClause);
                        break;
                    case DeclarationType.Parameter:
                        var parameterContext = (VBAParser.ArgContext)result.Target.Context;
                        rewriter.InsertAfter(parameterContext.unrestrictedIdentifier().Stop.TokenIndex, asTypeClause);
                        break;
                    case DeclarationType.Function:
                        var functionContext = (VBAParser.FunctionStmtContext) result.Target.Context;
                        rewriter.InsertAfter(functionContext.argList().Stop.TokenIndex, asTypeClause);
                        break;
                    case DeclarationType.PropertyGet:
                        var propertyContext = (VBAParser.PropertyGetStmtContext)result.Target.Context;
                        rewriter.InsertAfter(propertyContext.argList().Stop.TokenIndex, asTypeClause);
                        break;
                }
            }

            foreach (var reference in result.Target.References)
            {
                var context = reference.Context.GetDescendent<VBAParser.TypeHintContext>();

                if (context != null)
                {
                    var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                    rewriter.Remove(context);
                }
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveTypeHintsQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}