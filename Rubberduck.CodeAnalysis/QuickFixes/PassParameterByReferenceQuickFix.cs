using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class PassParameterByReferenceQuickFix : QuickFixBase
    {
        public PassParameterByReferenceQuickFix()
            : base(typeof(AssignedByValParameterInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);

            var token = ((VBAParser.ArgContext)result.Target.Context).BYVAL().Symbol;
            rewriter.Replace(token, Tokens.ByRef);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.PassParameterByReferenceQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}