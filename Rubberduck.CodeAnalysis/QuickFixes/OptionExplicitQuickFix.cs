using System;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class OptionExplicitQuickFix : QuickFixBase
    {
        public OptionExplicitQuickFix()
            : base(typeof(OptionExplicitInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(0, Tokens.Option + ' ' + Tokens.Explicit + Environment.NewLine + Environment.NewLine);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.OptionExplicitQuickFix;
        

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => true;
    }
}