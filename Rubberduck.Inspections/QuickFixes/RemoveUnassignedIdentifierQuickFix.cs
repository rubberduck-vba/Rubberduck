using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveUnassignedIdentifierQuickFix : QuickFixBase
    {
        private readonly Declaration _target;
        private readonly IModuleRewriter _rewriter;

        public RemoveUnassignedIdentifierQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, IModuleRewriter rewriter)
            : base(context, selection, InspectionsUI.RemoveUnassignedIdentifierQuickFix)
        {
            _target = target;
            _rewriter = rewriter;
        }

        public override void Fix()
        {
            _rewriter.Remove(_target);
        }
    }
}