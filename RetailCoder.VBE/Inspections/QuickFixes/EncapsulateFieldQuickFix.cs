using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// A code inspection quickfix that encapsulates a public field with a property
    /// </summary>
    public class EncapsulateFieldQuickFix : QuickFixBase
    {
        private readonly Declaration _target;
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;

        public EncapsulateFieldQuickFix(ParserRuleContext context, QualifiedSelection selection, Declaration target, RubberduckParserState state, IIndenter indenter)
            : base(context, selection, string.Format(InspectionsUI.EncapsulatePublicFieldInspectionQuickFix, target.IdentifierName))
        {
            _target = target;
            _state = state;
            _indenter = indenter;
        }

        public override void Fix()
        {
            var vbe = _target.Project.VBE;

            using (var view = new EncapsulateFieldDialog(_state, _indenter))
            {
                var factory = new EncapsulateFieldPresenterFactory(vbe, _state, view);
                var refactoring = new EncapsulateFieldRefactoring(vbe, _indenter, factory);
                refactoring.Refactor(_target);
                IsCancelled = view.DialogResult != DialogResult.OK;
            }
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }
    }
}