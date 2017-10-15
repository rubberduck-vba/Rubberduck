using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
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

        public RemoveUnusedDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target)
            : base(context, selection, InspectionsUI.RemoveUnusedDeclarationQuickFix)
        {
            _target = target;
        }

        public override void Fix()
        {
            var module = _target
                .QualifiedName
                .QualifiedModuleName
                .Component
                .CodeModule;

            module.Remove(_target);
        }
    }
}