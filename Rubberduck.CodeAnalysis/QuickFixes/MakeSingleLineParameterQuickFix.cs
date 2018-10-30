using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class MakeSingleLineParameterQuickFix : QuickFixBase
    {
        public MakeSingleLineParameterQuickFix()
            : base(typeof(MultilineParameterInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            var parameter = result.Context.GetText()
                .Replace("_", "")
                .RemoveExtraSpacesLeavingIndentation();

            rewriter.Replace(result.Context, parameter);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.MakeSingleLineParameterQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
