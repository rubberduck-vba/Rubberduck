using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// A code inspection quickfix that removes an unused identifier declaration.
    /// </summary>
    public class RemoveUnusedDeclarationQuickFix : QuickFixBase
    {
        private readonly Declaration _target;
        private readonly IModuleRewriter _rewriter;

        public RemoveUnusedDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, IModuleRewriter rewriter)
            : base(context, selection, InspectionsUI.RemoveUnusedDeclarationQuickFix)
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