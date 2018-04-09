using System;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class OptionExplicitQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public OptionExplicitQuickFix(RubberduckParserState state)
            : base(typeof(OptionExplicitInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.InsertBefore(0, Tokens.Option + ' ' + Tokens.Explicit + Environment.NewLine + Environment.NewLine);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.OptionExplicitQuickFix;
        

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => true;
    }
}