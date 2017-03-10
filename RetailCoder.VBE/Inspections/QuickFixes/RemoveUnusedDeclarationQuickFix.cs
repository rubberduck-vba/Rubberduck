using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
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
        private readonly TokenStreamRewriter _rewriter;

        public RemoveUnusedDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, TokenStreamRewriter rewriter)
            : base(context, selection, InspectionsUI.RemoveUnusedDeclarationQuickFix)
        {
            _target = target;
            _rewriter = rewriter;
        }

        public override void Fix()
        {
            var module = _target
                .QualifiedName
                .QualifiedModuleName
                .Component
                .CodeModule;

            module.Remove(_rewriter, _target);
        }
    }
}