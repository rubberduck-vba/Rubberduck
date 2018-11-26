using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class DeclareAsExplicitVariantQuickFix : QuickFixBase
    {
        public DeclareAsExplicitVariantQuickFix()
            : base(typeof(VariableTypeNotDeclaredInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);

            ParserRuleContext identifierNode =
                result.Context is VBAParser.VariableSubStmtContext || result.Context is VBAParser.ConstSubStmtContext
                ? result.Context.children[0]
                : ((dynamic) result.Context).unrestrictedIdentifier();
            rewriter.InsertAfter(identifierNode.Stop.TokenIndex, " As Variant");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.DeclareAsExplicitVariantQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}