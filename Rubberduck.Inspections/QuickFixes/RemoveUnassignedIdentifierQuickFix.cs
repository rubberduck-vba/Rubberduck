using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveUnassignedIdentifierQuickFix : IQuickFix
    {
        private readonly Declaration _target;
        private readonly IModuleRewriter _rewriter;

        public RemoveUnassignedIdentifierQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, IModuleRewriter rewriter)
            : base(context, selection, InspectionsUI.RemoveUnassignedIdentifierQuickFix)
        {
            _target = target;
            _rewriter = rewriter;
        }

        public void Fix(IInspectionResult result)
        {
            _rewriter.Remove(_target);
        }
    }
}