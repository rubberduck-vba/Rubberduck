using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveExplicitLetStatementQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public RemoveExplicitLetStatementQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ObsoleteLetStatementInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var context = (VBAParser.LetStmtContext) result.Context;
            rewriter.Remove(context.LET());
            rewriter.Remove(context.whiteSpace().First());
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RemoveObsoleteStatementQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}